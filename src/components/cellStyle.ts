import type { CSSProperties } from 'react';
import type { BorderSide, CellFormat, Value } from '../model/types';
import { alignFor, codeColor } from '../engine/format';

// How a cell's format is drawn on Cells' dark tiles. The stored format is never changed:
// colors from an Excel file (made for white paper) are only adjusted on screen when they
// would be unreadable, so they go back to Excel exactly as they came.

const TILE = '#222d3a';
const TEXT = '#dfe6ee';
const LINE = '#c9d2dc';

function luminance(hex: string) {
  const ch = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const [r, g, b] = ch.map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function contrast(a: string, b: string) {
  const [x, y] = [luminance(a), luminance(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}

const readableOn = (bg: string) => (luminance(bg) > 0.4 ? '#1b232d' : '#ffffff');

const BORDER_WIDTH = { hair: 1, thin: 1, medium: 2, thick: 3, double: 3, dashed: 1, dotted: 1 } as const;

function borderCss(b: BorderSide | undefined, bg: string) {
  if (!b) return undefined;
  const style = b.style === 'double' ? 'double' : b.style === 'dashed' ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid';
  const color = b.color && contrast(b.color, bg) >= 2 ? b.color : LINE;
  return `${BORDER_WIDTH[b.style]}px ${style} ${color}`;
}

export function cellLook(v: Value, fmt: CellFormat | undefined): CSSProperties {
  const align = alignFor(v, fmt);
  const s: CSSProperties = {
    justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
  };
  if (!fmt) return s;
  const bg = fmt.fill ?? TILE;
  if (fmt.fill) s.background = fmt.fill;

  const wanted = codeColor(v, fmt.numFmtCode) ?? fmt.color;
  if (fmt.fill) s.color = wanted && contrast(wanted, bg) >= 3 ? wanted : readableOn(bg);
  else if (wanted && contrast(wanted, TILE) >= 3) s.color = wanted;
  else if (wanted) s.color = TEXT;

  if (fmt.bold) s.fontWeight = 700;
  if (fmt.italic) s.fontStyle = 'italic';
  const deco = [fmt.underline && 'underline', fmt.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) s.textDecoration = deco;
  if (fmt.fontSize) s.fontSize = `${Math.round((fmt.fontSize * 13) / 11)}px`;
  if (fmt.fontFamily) s.fontFamily = `'${fmt.fontFamily}', var(--sans)`;
  if (fmt.valign) s.alignItems = fmt.valign === 'top' ? 'flex-start' : fmt.valign === 'bottom' ? 'flex-end' : 'center';
  if (fmt.wrap) {
    s.whiteSpace = 'normal';
    s.lineHeight = 1.25;
    s.textAlign = align;
    s.paddingTop = 4;
    s.paddingBottom = 4;
  }
  if (fmt.border) {
    s.borderTop = borderCss(fmt.border.top, bg);
    s.borderRight = borderCss(fmt.border.right, bg);
    s.borderBottom = borderCss(fmt.border.bottom, bg);
    s.borderLeft = borderCss(fmt.border.left, bg);
  }
  return s;
}
