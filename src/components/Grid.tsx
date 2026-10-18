import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent, type MouseEvent as ReactMouseEvent } from 'react';
import { colName, inRect, rectOf, toKey, type Pos, type Rect } from '../model/address';
import type { Sheet, Value } from '../model/types';
import { isError } from '../model/types';
import { formatValue } from '../engine/format';
import { cellLook, richRunLook } from './cellStyle';
import { mergedBorder } from '../model/merge';

export interface Selection {
  anchor: Pos;
  focus: Pos;
}

export interface GridActions {
  select(sel: Selection): void;
  startEdit(initial?: string): void;
  commitEdit(raw: string, move: 'down' | 'right' | 'none'): void;
  cancelEdit(): void;
  clear(): void;
  undo(): void;
  redo(): void;
  toggleBold(): void;
  addRows(): void;
  resizeColumn(c: number, px: number): void;
}

interface Props {
  sheet: Sheet;
  values: Record<string, Value>;
  /** A1 keys showing a value Excel calculated because Cells can't run the formula yet. */
  approx: Set<string>;
  sel: Selection;
  editing: { initial: string } | null;
  show: { colHeads: boolean; rowHeads: boolean; formulas: boolean; compact: boolean };
  actions: GridActions;
  /** Cells where other people have their cursor. */
  peers: { key: string; name: string; color: string }[];
  /** The workbook's base font, for files that bring their own (Excel's "Normal" style). */
  baseFont?: { name: string; size: number };
  onCopy(cut: boolean): string;
  onPaste(text: string): void;
}

export const DEFAULT_COL_W = 112;
const HEAD_W = 40;
const HEAD_H = 26;
const GAP = { normal: 6, compact: 3 };
const ROW_H = { normal: 32, compact: 26 };

const clamp = (n: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, n));

export function Grid({ sheet, values, approx, sel, editing, show, actions, peers, baseFont, onCopy, onPaste }: Props) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const dragging = useRef(false);
  const [resizing, setResizing] = useState<{ c: number; px: number } | null>(null);
  const rect = rectOf(sel.anchor, sel.focus);
  const { r: ar, c: ac } = sel.focus;
  const peerAt = new Map(peers.map((p) => [p.key, p]));

  const layout = sheet.layout ?? {};
  const gap = show.compact ? GAP.compact : GAP.normal;
  const defaultRowH = show.compact ? ROW_H.compact : ROW_H.normal;
  const hiddenCols = new Set(layout.hiddenCols ?? []);
  const hiddenRows = new Set(layout.hiddenRows ?? []);
  const colW = (c: number) =>
    hiddenCols.has(c) ? 0 : resizing?.c === c ? resizing.px : layout.colWidths?.[c] ?? DEFAULT_COL_W;
  const rowH = (r: number) => (hiddenRows.has(r) ? 0 : layout.rowHeights?.[r] ?? defaultRowH);

  // Merged blocks: the top-left cell spans the block, the others aren't drawn.
  const merges = layout.merges ?? [];
  const mergeAt = (r: number, c: number): Rect | undefined => merges.find((m) => inRect(m, r, c));

  useEffect(() => {
    const up = () => (dragging.current = false);
    window.addEventListener('mouseup', up);
    return () => window.removeEventListener('mouseup', up);
  }, []);

  useEffect(() => {
    scrollRef.current
      ?.querySelector<HTMLElement>(`[data-key="${toKey(ar, ac)}"]`)
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [ar, ac]);

  // Keep keyboard focus on the grid after an edit so typing keeps working.
  useEffect(() => {
    if (!editing) scrollRef.current?.focus({ preventScroll: true });
  }, [editing]);

  function move(dr: number, dc: number, extend: boolean) {
    let r = sel.focus.r;
    let c = sel.focus.c;
    // Step over hidden rows/columns and the covered part of merged blocks, like Excel.
    const from = mergeAt(r, c);
    if (from) {
      if (dr > 0) r = from.r1;
      if (dc > 0) c = from.c1;
      if (dr < 0) r = from.r0;
      if (dc < 0) c = from.c0;
    }
    do {
      r = clamp(r + dr, 0, sheet.rows - 1);
      c = clamp(c + dc, 0, sheet.cols - 1);
    } while ((hiddenRows.has(r) && dr && r > 0 && r < sheet.rows - 1) || (hiddenCols.has(c) && dc && c > 0 && c < sheet.cols - 1));
    const m = mergeAt(r, c);
    const focus = m ? { r: m.r0, c: m.c0 } : { r, c };
    actions.select({ anchor: extend ? sel.anchor : focus, focus });
  }

  function onKeyDown(e: KeyboardEvent<HTMLDivElement>) {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod) {
      const k = e.key.toLowerCase();
      if (k === 'z') { e.preventDefault(); if (e.shiftKey) actions.redo(); else actions.undo(); }
      else if (k === 'y') { e.preventDefault(); actions.redo(); }
      else if (k === 'b' || k === 'n') { e.preventDefault(); actions.toggleBold(); }
      else if (k === 'a') { e.preventDefault(); actions.select({ anchor: { r: 0, c: 0 }, focus: { r: sheet.rows - 1, c: sheet.cols - 1 } }); }
      return;
    }
    switch (e.key) {
      case 'ArrowUp': e.preventDefault(); return move(-1, 0, e.shiftKey);
      case 'ArrowDown': e.preventDefault(); return move(1, 0, e.shiftKey);
      case 'ArrowLeft': e.preventDefault(); return move(0, -1, e.shiftKey);
      case 'ArrowRight': e.preventDefault(); return move(0, 1, e.shiftKey);
      case 'Tab': e.preventDefault(); return move(0, e.shiftKey ? -1 : 1, false);
      case 'Enter':
      case 'F2': e.preventDefault(); return actions.startEdit();
      case 'Delete':
      case 'Backspace': e.preventDefault(); return actions.clear();
    }
    if (e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      actions.startEdit(e.key);
    }
  }

  function startResize(e: ReactMouseEvent, c: number) {
    e.preventDefault();
    e.stopPropagation();
    const x0 = e.clientX;
    const w0 = colW(c);
    let px = w0;
    const onMove = (ev: MouseEvent) => {
      px = Math.max(24, Math.round(w0 + ev.clientX - x0));
      setResizing({ c, px });
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      setResizing(null);
      if (px !== w0) actions.resizeColumn(c, px);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }

  const cols = Array.from({ length: sheet.cols }, (_, c) => c);
  const rows = Array.from({ length: sheet.rows }, (_, r) => r);
  const rowOffset = show.colHeads ? 2 : 1; // CSS grid lines are 1-based; headers take the first track
  const colOffset = show.rowHeads ? 2 : 1;

  // Frozen panes: those rows/columns stick while the rest scrolls.
  const frozen = layout.frozen ?? { rows: 0, cols: 0 };
  const stickTop: number[] = [];
  for (let r = 0, y = show.colHeads ? HEAD_H + gap : 0; r < frozen.rows; r++) {
    stickTop.push(y);
    y += rowH(r) + gap;
  }
  const stickLeft: number[] = [];
  for (let c = 0, x = show.rowHeads ? HEAD_W + gap : 0; c < frozen.cols; c++) {
    stickLeft.push(x);
    x += colW(c) + gap;
  }

  return (
    <div
      ref={scrollRef}
      className="grid-scroll"
      tabIndex={0}
      role="grid"
      aria-label={`Hoja ${sheet.name}`}
      aria-rowcount={sheet.rows}
      aria-colcount={sheet.cols}
      onKeyDown={onKeyDown}
      onCopy={(e) => { if (!editing) { e.preventDefault(); e.clipboardData.setData('text/plain', onCopy(false)); } }}
      onCut={(e) => { if (!editing) { e.preventDefault(); e.clipboardData.setData('text/plain', onCopy(true)); } }}
      onPaste={(e) => { if (!editing) { e.preventDefault(); onPaste(e.clipboardData.getData('text/plain')); } }}
    >
      <div
        className="grid"
        style={{
          gap,
          ...(baseFont
            ? ({ '--cell-font': `'${baseFont.name}', var(--sans)`, '--cell-size': `${Math.round((baseFont.size * 13) / 11)}px` } as CSSProperties)
            : {}),
          gridTemplateColumns: `${show.rowHeads ? `${HEAD_W}px ` : ''}${cols.map((c) => `${colW(c)}px`).join(' ')}`,
          gridTemplateRows: `${show.colHeads ? `${HEAD_H}px ` : ''}${rows.map((r) => `${rowH(r)}px`).join(' ')}`,
        }}
      >
        {show.colHeads && (
          <div className="grid-row" role="row">
            {show.rowHeads && <div className="corner" style={{ gridRow: 1, gridColumn: 1 }} />}
            {cols.map((c) =>
              hiddenCols.has(c) ? null : (
                <div
                  key={c}
                  role="columnheader"
                  className={`col-head${c >= rect.c0 && c <= rect.c1 ? ' on' : ''}`}
                  style={{ gridRow: 1, gridColumn: c + colOffset, ...(c < frozen.cols ? { left: stickLeft[c], zIndex: 4 } : {}) }}
                >
                  {colName(c)}
                  <span className="col-resize" aria-hidden="true" onMouseDown={(e) => startResize(e, c)} />
                </div>
              ),
            )}
          </div>
        )}
        {rows.map((r) =>
          hiddenRows.has(r) ? null : (
            <div key={r} className="grid-row" role="row">
              {show.rowHeads && (
                <div
                  role="rowheader"
                  className={`row-head${r >= rect.r0 && r <= rect.r1 ? ' on' : ''}`}
                  style={{ gridRow: r + rowOffset, gridColumn: 1, ...(r < frozen.rows ? { top: stickTop[r], zIndex: 4 } : {}) }}
                >
                  {r + 1}
                </div>
              )}
              {cols.map((c) => {
                if (hiddenCols.has(c)) return null;
                const merge = mergeAt(r, c);
                if (merge && (merge.r0 !== r || merge.c0 !== c)) return null;
                const key = toKey(r, c);
                const cell = sheet.cells[key];
                const active = r === ar && c === ac;
                const v = values[key] ?? null;
                // A merged block is drawn as one cell, with the outline gathered from its edges.
                const fmt = merge && cell ? { ...cell.fmt, border: mergedBorder(sheet, merge) } : cell?.fmt;
                const text = show.formulas && cell?.raw.startsWith('=') ? cell.raw : formatValue(v, fmt);
                const cls = ['cell'];
                if (fmt?.tint && fmt.tint !== 'none' && !fmt.fill) cls.push(`tint-${fmt.tint}`);
                if (typeof v === 'number') cls.push('num');
                if (isError(v)) cls.push('err');
                if (inRect(rect, r, c) && !(rect.r0 === rect.r1 && rect.c0 === rect.c1)) cls.push('in-sel');
                if (active) cls.push('active');
                if (approx.has(key)) cls.push('approx');
                if (cell?.link) cls.push('link');
                const peer = peerAt.get(key);
                const sticky: CSSProperties = {};
                if (r < frozen.rows) sticky.top = stickTop[r];
                if (c < frozen.cols) sticky.left = stickLeft[c];
                if (r < frozen.rows || c < frozen.cols) {
                  sticky.position = 'sticky';
                  sticky.zIndex = r < frozen.rows && c < frozen.cols ? 3 : 2;
                }
                const shadows = [
                  peer && !active ? `inset 0 0 0 2px ${peer.color}` : '',
                  sticky.position ? `0 0 0 ${gap}px var(--well)` : '',
                ].filter(Boolean);
                const title = [
                  isError(v) ? errorHelp(v.error) : '',
                  approx.has(key) ? 'Cells todavía no calcula esta fórmula: muestra el último resultado guardado por Excel o Google Sheets.' : '',
                  cell?.note ?? '',
                  cell?.link ? `${cell.link} (Ctrl+clic para abrir)` : '',
                ].filter(Boolean).join('\n');
                return (
                  <div
                    key={c}
                    role="gridcell"
                    aria-selected={inRect(rect, r, c)}
                    data-key={key}
                    className={cls.join(' ')}
                    style={{
                      width: 'auto',
                      height: 'auto',
                      gridRow: merge ? `${r + rowOffset} / span ${merge.r1 - merge.r0 + 1}` : r + rowOffset,
                      gridColumn: merge ? `${c + colOffset} / span ${merge.c1 - merge.c0 + 1}` : c + colOffset,
                      ...cellLook(v, fmt),
                      ...sticky,
                      ...(shadows.length && !active ? { boxShadow: shadows.join(', ') } : {}),
                    }}
                    title={title || undefined}
                    onMouseDown={(e) => {
                      if (e.button !== 0) return;
                      if (cell?.link && (e.ctrlKey || e.metaKey)) {
                        window.open(cell.link, '_blank', 'noopener,noreferrer');
                        return;
                      }
                      if (editing) return; // the editor's blur commits; then this click selects
                      dragging.current = true;
                      actions.select({ anchor: e.shiftKey ? sel.anchor : { r, c }, focus: { r, c } });
                    }}
                    onMouseEnter={() => {
                      if (dragging.current) actions.select({ anchor: sel.anchor, focus: { r, c } });
                    }}
                    onDoubleClick={() => actions.startEdit()}
                  >
                    {active && editing ? (
                      <CellEditor initial={editing.initial} options={cell?.validation?.list ?? null} actions={actions} />
                    ) : cell?.rich && !show.formulas ? (
                      <span>
                        {cell.rich.map((run, i) => (
                          <span key={i} style={richRunLook(run, fmt)}>
                            {run.text}
                          </span>
                        ))}
                      </span>
                    ) : (
                      text
                    )}
                    {cell?.note && <span className="note-mark" aria-label="Tiene una nota" />}
                    {peer && (
                      <span className="peer-tag" style={{ background: peer.color }}>
                        {peer.name}
                      </span>
                    )}
                  </div>
                );
              })}
            </div>
          ),
        )}
      </div>
      <div className="grid-more">
        <button type="button" className="pill" onClick={actions.addRows}>
          + 50 filas
        </button>
      </div>
    </div>
  );
}

function CellEditor({ initial, options, actions }: { initial: string; options: string[] | null; actions: GridActions }) {
  const [text, setText] = useState(initial);
  const done = useRef(false);
  const finish = (move: 'down' | 'right' | 'none') => {
    if (done.current) return;
    done.current = true;
    actions.commitEdit(text, move);
  };
  return (
    <>
    {options && (
      <datalist id="cell-options">
        {options.map((o) => (
          <option key={o} value={o} />
        ))}
      </datalist>
    )}
    <input
      autoFocus
      list={options ? 'cell-options' : undefined}
      aria-label="Editar celda"
      value={text}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => finish('none')}
      onMouseDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.key === 'Enter') { e.preventDefault(); finish('down'); }
        else if (e.key === 'Tab') { e.preventDefault(); finish('right'); }
        else if (e.key === 'Escape') { e.preventDefault(); done.current = true; actions.cancelEdit(); }
      }}
    />
    </>
  );
}

function errorHelp(code: string): string {
  switch (code) {
    case '#DIV/0!': return 'División por cero';
    case '#CIRC!': return 'Referencia circular: la fórmula depende de sí misma';
    case '#REF!': return 'Referencia inválida (¿existe esa hoja?)';
    case '#NAME?': return 'Nombre de función desconocido';
    case '#VALUE!': return 'Tipo de valor incorrecto';
    case '#ERROR!': return 'No se pudo leer la fórmula';
    default: return code;
  }
}
