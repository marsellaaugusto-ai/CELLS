import * as Y from 'yjs';
import type { Cell, Sheet, Workbook } from '../model/types';
import { applyEdit, type Edit } from '../state/workbook';

// A workbook lives in a Yjs document so several people can edit it at once:
//   meta:   Y.Map  { id, name, updatedAt }
//   sheets: Y.Array of Y.Map { id, name, rows, cols, cells: Y.Map<A1 key, Cell> }
// Each cell is replaced as a whole, so edits to different cells always merge and
// simultaneous edits to the same cell end with one of them, the same on every screen.

/** Origin of changes that came from another person; they are not sent back out or undone. */
export const REMOTE = Symbol('remote');

type YSheet = Y.Map<unknown>;

function sheetToY(s: Sheet): YSheet {
  const y = new Y.Map<unknown>();
  y.set('id', s.id);
  y.set('name', s.name);
  y.set('rows', s.rows);
  y.set('cols', s.cols);
  const cells = new Y.Map<Cell>();
  for (const [k, c] of Object.entries(s.cells)) cells.set(k, c);
  y.set('cells', cells);
  return y;
}

function sheetFromY(y: YSheet): Sheet {
  return {
    id: y.get('id') as string,
    name: y.get('name') as string,
    rows: y.get('rows') as number,
    cols: y.get('cols') as number,
    cells: (y.get('cells') as Y.Map<Cell>).toJSON() as Record<string, Cell>,
  };
}

export function workbookFromDoc(doc: Y.Doc): Workbook {
  const meta = doc.getMap<string>('meta');
  return {
    id: meta.get('id') ?? '',
    name: meta.get('name') ?? '',
    updatedAt: meta.get('updatedAt') ?? '',
    sheets: doc.getArray<YSheet>('sheets').toArray().map(sheetFromY),
  };
}

/**
 * Builds the document for a workbook saved before collaboration existed.
 * It is written with a fixed client id, so two people converting the same saved
 * workbook produce identical history and their copies merge instead of duplicating.
 */
export function docStateFromWorkbook(wb: Workbook): Uint8Array {
  const doc = new Y.Doc();
  doc.clientID = 0;
  doc.transact(() => {
    const meta = doc.getMap<string>('meta');
    meta.set('id', wb.id);
    meta.set('name', wb.name);
    meta.set('updatedAt', wb.updatedAt);
    doc.getArray<YSheet>('sheets').push(wb.sheets.map(sheetToY));
  });
  const state = Y.encodeStateAsUpdate(doc);
  doc.destroy();
  return state;
}

/** Writes into the document only what differs between two snapshots. */
function writeDiff(doc: Y.Doc, before: Workbook, after: Workbook) {
  const meta = doc.getMap<string>('meta');
  if (before.name !== after.name) meta.set('name', after.name);

  const ySheets = doc.getArray<YSheet>('sheets');
  const keep = new Set(after.sheets.map((s) => s.id));
  for (let i = ySheets.length - 1; i >= 0; i--) {
    if (!keep.has(ySheets.get(i).get('id') as string)) ySheets.delete(i, 1);
  }
  const byId = new Map(ySheets.toArray().map((y) => [y.get('id') as string, y]));
  for (const s of after.sheets) {
    const y = byId.get(s.id);
    if (!y) {
      ySheets.push([sheetToY(s)]);
      continue;
    }
    const b = before.sheets.find((x) => x.id === s.id);
    if (!b || b === s) continue;
    if (b.name !== s.name) y.set('name', s.name);
    if (b.rows !== s.rows) y.set('rows', Math.max(s.rows, y.get('rows') as number));
    if (b.cols !== s.cols) y.set('cols', Math.max(s.cols, y.get('cols') as number));
    if (b.cells === s.cells) continue;
    const yCells = y.get('cells') as Y.Map<Cell>;
    for (const k of new Set([...Object.keys(b.cells), ...Object.keys(s.cells)])) {
      if (b.cells[k] === s.cells[k]) continue;
      if (s.cells[k]) yCells.set(k, s.cells[k]);
      else yCells.delete(k);
    }
  }
}

/** One open workbook: its shared document, a snapshot for the UI, and per-person undo. */
export class Book {
  readonly doc = new Y.Doc();
  private snap: Workbook;
  private listeners = new Set<() => void>();
  private undoManager: Y.UndoManager;

  constructor(state: Uint8Array) {
    Y.applyUpdate(this.doc, state, REMOTE);
    this.snap = workbookFromDoc(this.doc);
    // Undo only reverts this person's own changes, never someone else's.
    this.undoManager = new Y.UndoManager([this.doc.getMap('meta'), this.doc.getArray('sheets')], {
      captureTimeout: 300,
      trackedOrigins: new Set([this]),
    });
    this.doc.on('update', () => {
      this.snap = workbookFromDoc(this.doc);
      this.listeners.forEach((fn) => fn());
    });
  }

  static fromWorkbook(wb: Workbook) {
    return new Book(docStateFromWorkbook(wb));
  }

  get snapshot() {
    return this.snap;
  }

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  edit(e: Edit) {
    const before = this.snap;
    const after = applyEdit(before, e);
    if (after === before) return;
    this.doc.transact(() => {
      writeDiff(this.doc, before, after);
      this.doc.getMap<string>('meta').set('updatedAt', new Date().toISOString());
    }, this);
  }

  undo() {
    this.undoManager.undo();
  }

  redo() {
    this.undoManager.redo();
  }

  /** True for changes made here (edits, undo, redo), false for ones received from others. */
  isLocal(origin: unknown) {
    return origin === this || origin === this.undoManager;
  }

  encode() {
    return Y.encodeStateAsUpdate(this.doc);
  }

  merge(state: Uint8Array) {
    Y.applyUpdate(this.doc, state, REMOTE);
  }

  destroy() {
    this.undoManager.destroy();
    this.doc.destroy();
    this.listeners.clear();
  }
}

export function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function fromBase64(b64: string): Uint8Array {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
