import { parseKey, rectKeys, toKey, type Rect } from '../model/address';

const overlaps = (a: Rect, b: Rect) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
import type { Cell, CellFormat, Sheet, SheetLayout, Workbook } from '../model/types';
import { newSheet } from '../model/types';

export type Edit =
  | { type: 'setCell'; sheetId: string; key: string; raw: string }
  | { type: 'format'; sheetId: string; rect: Rect; patch: CellFormat }
  | { type: 'clear'; sheetId: string; rect: Rect }
  | { type: 'paste'; sheetId: string; r: number; c: number; rows: string[][] }
  | { type: 'sortRows'; sheetId: string; rect: Rect; byCol: number; desc: boolean }
  | { type: 'addRows'; sheetId: string; count: number }
  | { type: 'addSheet' }
  | { type: 'renameSheet'; sheetId: string; name: string }
  | { type: 'deleteSheet'; sheetId: string }
  | { type: 'renameBook'; name: string }
  | { type: 'layout'; sheetId: string; patch: Partial<SheetLayout> }
  | { type: 'merge'; sheetId: string; rect: Rect }
  | { type: 'unmerge'; sheetId: string; rect: Rect };

function withCell(cells: Record<string, Cell>, key: string, cell: Cell | undefined) {
  const next = { ...cells };
  if (!cell || (cell.raw === '' && !cell.fmt)) delete next[key];
  else next[key] = cell;
  return next;
}

function updateSheet(wb: Workbook, sheetId: string, fn: (s: Sheet) => Sheet): Workbook {
  return { ...wb, sheets: wb.sheets.map((s) => (s.id === sheetId ? fn(s) : s)) };
}

function uniqueSheetName(wb: Workbook) {
  let n = wb.sheets.length + 1;
  while (wb.sheets.some((s) => s.name === `Hoja ${n}`)) n++;
  return `Hoja ${n}`;
}

export function applyEdit(wb: Workbook, edit: Edit): Workbook {
  switch (edit.type) {
    case 'setCell':
      return updateSheet(wb, edit.sheetId, (s) => {
        const prev = s.cells[edit.key];
        return { ...s, cells: withCell(s.cells, edit.key, { ...prev, raw: edit.raw }) };
      });
    case 'format':
      return updateSheet(wb, edit.sheetId, (s) => {
        let cells = s.cells;
        for (const key of rectKeys(edit.rect)) {
          const prev = cells[key] ?? { raw: '' };
          cells = withCell(cells, key, { ...prev, fmt: { ...prev.fmt, ...edit.patch } });
        }
        return { ...s, cells };
      });
    case 'clear':
      return updateSheet(wb, edit.sheetId, (s) => {
        const cells = { ...s.cells };
        for (const key of rectKeys(edit.rect)) {
          if (!cells[key]) continue;
          if (cells[key].fmt) cells[key] = { ...cells[key], raw: '' };
          else delete cells[key];
        }
        return { ...s, cells };
      });
    case 'paste':
      return updateSheet(wb, edit.sheetId, (s) => {
        let cells = s.cells;
        edit.rows.forEach((row, i) =>
          row.forEach((raw, j) => {
            const key = toKey(edit.r + i, edit.c + j);
            cells = withCell(cells, key, { ...cells[key], raw });
          }),
        );
        return {
          ...s,
          cells,
          rows: Math.max(s.rows, edit.r + edit.rows.length),
          cols: Math.max(s.cols, edit.c + Math.max(0, ...edit.rows.map((r) => r.length))),
        };
      });
    case 'sortRows':
      return updateSheet(wb, edit.sheetId, (s) => {
        const { r0, r1, c0, c1 } = edit.rect;
        const rows: (Cell | undefined)[][] = [];
        for (let r = r0; r <= r1; r++) {
          const row: (Cell | undefined)[] = [];
          for (let c = c0; c <= c1; c++) row.push(s.cells[toKey(r, c)]);
          rows.push(row);
        }
        const idx = edit.byCol - c0;
        const coll = new Intl.Collator('es', { numeric: true, sensitivity: 'base' });
        rows.sort((a, b) => {
          const x = a[idx]?.raw ?? '';
          const y = b[idx]?.raw ?? '';
          if (x === '' || y === '') return x === y ? 0 : x === '' ? 1 : -1; // blanks last
          const nx = Number(x.replace(',', '.'));
          const ny = Number(y.replace(',', '.'));
          const cmp = !isNaN(nx) && !isNaN(ny) ? nx - ny : coll.compare(x, y);
          return edit.desc ? -cmp : cmp;
        });
        let cells = s.cells;
        rows.forEach((row, i) => row.forEach((cell, j) => (cells = withCell(cells, toKey(r0 + i, c0 + j), cell))));
        return { ...s, cells };
      });
    case 'addRows':
      return updateSheet(wb, edit.sheetId, (s) => ({ ...s, rows: s.rows + edit.count }));
    case 'addSheet':
      return { ...wb, sheets: [...wb.sheets, newSheet(uniqueSheetName(wb))] };
    case 'renameSheet': {
      const name = edit.name.trim();
      if (!name || wb.sheets.some((s) => s.id !== edit.sheetId && s.name.toLowerCase() === name.toLowerCase())) return wb;
      return updateSheet(wb, edit.sheetId, (s) => ({ ...s, name }));
    }
    case 'deleteSheet':
      return wb.sheets.length > 1 ? { ...wb, sheets: wb.sheets.filter((s) => s.id !== edit.sheetId) } : wb;
    case 'renameBook':
      return edit.name.trim() ? { ...wb, name: edit.name.trim() } : wb;
    case 'layout':
      return updateSheet(wb, edit.sheetId, (s) => ({ ...s, layout: { ...s.layout, ...edit.patch } }));
    case 'merge': {
      const { rect } = edit;
      if (rect.r0 === rect.r1 && rect.c0 === rect.c1) return wb;
      return updateSheet(wb, edit.sheetId, (s) => {
        // Like Excel, only the top-left value survives a merge.
        let cells = s.cells;
        for (const key of rectKeys(rect)) {
          if (key === toKey(rect.r0, rect.c0) || !cells[key]) continue;
          cells = withCell(cells, key, { ...cells[key], raw: '' });
        }
        const merges = (s.layout?.merges ?? []).filter((m) => !overlaps(m, rect));
        return { ...s, cells, layout: { ...s.layout, merges: [...merges, { ...rect }] } };
      });
    }
    case 'unmerge':
      return updateSheet(wb, edit.sheetId, (s) => ({
        ...s,
        layout: { ...s.layout, merges: (s.layout?.merges ?? []).filter((m) => !overlaps(m, edit.rect)) },
      }));
  }
}

/** Raw text of a block of cells, tab-separated like Excel and Google Sheets put on the clipboard. */
export function rectToTsv(sheet: Sheet, rect: Rect): string {
  const lines: string[] = [];
  for (let r = rect.r0; r <= rect.r1; r++) {
    const row: string[] = [];
    for (let c = rect.c0; c <= rect.c1; c++) row.push(sheet.cells[toKey(r, c)]?.raw ?? '');
    lines.push(row.join('\t'));
  }
  return lines.join('\n');
}

export function usedRect(sheet: Sheet): Rect | null {
  const keys = Object.keys(sheet.cells).filter((k) => sheet.cells[k].raw !== '');
  if (!keys.length) return null;
  let r1 = 0;
  let c1 = 0;
  for (const k of keys) {
    const { r, c } = parseKey(k);
    r1 = Math.max(r1, r);
    c1 = Math.max(c1, c);
  }
  return { r0: 0, c0: 0, r1, c1 };
}
