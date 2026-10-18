import { toKey } from './address';
import type { CellFormat, Range, Sheet } from './types';

/**
 * The outline of a merged block. Excel keeps each side's border on the cell at that edge
 * (the right border of A1:C1 lives on C1), but a merged block is drawn — and, with ExcelJS,
 * saved — as one cell, so the sides are gathered onto it.
 */
export function mergedBorder(sheet: Sheet, m: Range): CellFormat['border'] {
  const side = (r: number, c: number, s: 'top' | 'right' | 'bottom' | 'left') => sheet.cells[toKey(r, c)]?.fmt?.border?.[s];
  const first = <T,>(xs: (T | undefined)[]) => xs.find((x) => x !== undefined);
  const cols = Array.from({ length: m.c1 - m.c0 + 1 }, (_, i) => m.c0 + i);
  const rows = Array.from({ length: m.r1 - m.r0 + 1 }, (_, i) => m.r0 + i);
  const border = {
    top: first(cols.map((c) => side(m.r0, c, 'top'))),
    bottom: first(cols.map((c) => side(m.r1, c, 'bottom'))),
    left: first(rows.map((r) => side(r, m.c0, 'left'))),
    right: first(rows.map((r) => side(r, m.c1, 'right'))),
  };
  const defined = Object.fromEntries(Object.entries(border).filter(([, v]) => v));
  return Object.keys(defined).length ? defined : undefined;
}
