import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { colName, inRect, rectOf, toKey, type Pos } from '../model/address';
import type { Sheet, Value } from '../model/types';
import { isError } from '../model/types';
import { alignFor, formatValue } from '../engine/format';

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
}

interface Props {
  sheet: Sheet;
  values: Record<string, Value>;
  sel: Selection;
  editing: { initial: string } | null;
  show: { colHeads: boolean; rowHeads: boolean; formulas: boolean };
  actions: GridActions;
  /** Cells where other people have their cursor. */
  peers: { key: string; name: string; color: string }[];
  onCopy(cut: boolean): string;
  onPaste(text: string): void;
}

const clamp = (n: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, n));

export function Grid({ sheet, values, sel, editing, show, actions, peers, onCopy, onPaste }: Props) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const dragging = useRef(false);
  const rect = rectOf(sel.anchor, sel.focus);
  const { r: ar, c: ac } = sel.focus;
  const peerAt = new Map(peers.map((p) => [p.key, p]));

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
    const focus = { r: clamp(sel.focus.r + dr, 0, sheet.rows - 1), c: clamp(sel.focus.c + dc, 0, sheet.cols - 1) };
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

  const cols = Array.from({ length: sheet.cols }, (_, c) => c);
  const rows = Array.from({ length: sheet.rows }, (_, r) => r);
  const headW = show.rowHeads ? '40px ' : '';
  const headH = show.colHeads ? '26px ' : '';

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
          gridTemplateColumns: `${headW}repeat(${sheet.cols}, var(--col-w))`,
          gridTemplateRows: `${headH}repeat(${sheet.rows}, var(--row-h))`,
        }}
      >
        {show.colHeads && (
          <div className="grid-row" role="row">
            {show.rowHeads && <div className="corner" />}
            {cols.map((c) => (
              <div key={c} role="columnheader" className={`col-head${c >= rect.c0 && c <= rect.c1 ? ' on' : ''}`}>
                {colName(c)}
              </div>
            ))}
          </div>
        )}
        {rows.map((r) => (
          <div key={r} className="grid-row" role="row">
            {show.rowHeads && (
              <div role="rowheader" className={`row-head${r >= rect.r0 && r <= rect.r1 ? ' on' : ''}`}>
                {r + 1}
              </div>
            )}
            {cols.map((c) => {
              const key = toKey(r, c);
              const cell = sheet.cells[key];
              const active = r === ar && c === ac;
              const v = values[key] ?? null;
              const fmt = cell?.fmt;
              const text = show.formulas && cell?.raw.startsWith('=') ? cell.raw : formatValue(v, fmt);
              const cls = ['cell'];
              if (fmt?.tint && fmt.tint !== 'none') cls.push(`tint-${fmt.tint}`);
              if (typeof v === 'number') cls.push('num');
              if (isError(v)) cls.push('err');
              if (inRect(rect, r, c) && !(rect.r0 === rect.r1 && rect.c0 === rect.c1)) cls.push('in-sel');
              if (active) cls.push('active');
              const align = alignFor(v, fmt);
              const peer = peerAt.get(key);
              return (
                <div
                  key={c}
                  role="gridcell"
                  aria-selected={inRect(rect, r, c)}
                  data-key={key}
                  className={cls.join(' ')}
                  style={{
                    justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
                    fontWeight: fmt?.bold ? 700 : undefined,
                    fontStyle: fmt?.italic ? 'italic' : undefined,
                    boxShadow: peer && !active ? `inset 0 0 0 2px ${peer.color}` : undefined,
                  }}
                  title={isError(v) ? errorHelp(v.error) : undefined}
                  onMouseDown={(e) => {
                    if (e.button !== 0) return;
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
                    <CellEditor initial={editing.initial} actions={actions} />
                  ) : (
                    text
                  )}
                  {peer && (
                    <span className="peer-tag" style={{ background: peer.color }}>
                      {peer.name}
                    </span>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
      <div className="grid-more">
        <button type="button" className="pill" onClick={actions.addRows}>
          + 50 filas
        </button>
      </div>
    </div>
  );
}

function CellEditor({ initial, actions }: { initial: string; actions: GridActions }) {
  const [text, setText] = useState(initial);
  const done = useRef(false);
  const finish = (move: 'down' | 'right' | 'none') => {
    if (done.current) return;
    done.current = true;
    actions.commitEdit(text, move);
  };
  return (
    <input
      autoFocus
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
