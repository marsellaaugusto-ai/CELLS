import FormulaParser from 'fast-formula-parser';
import { parseKey, toKey } from '../model/address';
import type { Sheet, Value, Workbook } from '../model/types';
import { isError } from '../model/types';
import { translateFormula } from './translate';

const { FormulaError } = FormulaParser;

/** Computed values for every non-empty cell, by sheet id then A1 key. */
export type Computed = Record<string, Record<string, Value>>;

export function parseLiteral(raw: string): Value {
  const s = raw.trim();
  if (s === '') return null;
  const upper = s.toUpperCase();
  if (upper === 'TRUE' || upper === 'VERDADERO') return true;
  if (upper === 'FALSE' || upper === 'FALSO') return false;
  const pct = s.endsWith('%');
  const body = pct ? s.slice(0, -1).trim() : s;
  let n: number | null = null;
  if (/^-?\d+([.,]\d+)?$/.test(body)) n = Number(body.replace(',', '.'));
  else if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(body)) n = Number(body.replace(/\./g, '').replace(',', '.'));
  else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(body)) n = Number(body.replace(/,/g, ''));
  if (n !== null && Number.isFinite(n)) return pct ? n / 100 : n;
  return raw;
}

const isFormula = (raw: string) => raw.length > 1 && raw.startsWith('=');

function usedBounds(sheet: Sheet) {
  let maxR = -1;
  let maxC = -1;
  for (const key of Object.keys(sheet.cells)) {
    const { r, c } = parseKey(key);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { maxR, maxC };
}

function fromEngine(v: unknown): Value {
  if (v instanceof FormulaError) return { error: (v as { error: string }).error };
  if (Array.isArray(v)) return fromEngine(Array.isArray(v[0]) ? v[0][0] : v[0]);
  if (v === undefined || v === null) return 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : { error: '#NUM!' };
  if (typeof v === 'string' || typeof v === 'boolean') return v;
  return { error: '#VALUE!' };
}

function toEngine(v: Value): unknown {
  return isError(v) ? new FormulaError(v.error) : v;
}

export function evaluateWorkbook(wb: Workbook): Computed {
  const byName = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s]));
  const bounds = new Map(wb.sheets.map((s) => [s.id, usedBounds(s)]));
  const cache = new Map<string, Value>();
  const visiting = new Set<string>();

  const sheetNamed = (name: string) => byName.get(String(name).toLowerCase());

  const parser = new FormulaParser({
    onCell: ({ sheet, row, col }: { sheet: string; row: number; col: number }) => {
      const s = sheetNamed(sheet);
      if (!s) return new FormulaError('#REF!');
      return toEngine(valueOf(s, toKey(row - 1, col - 1)));
    },
    onRange: (ref: { sheet: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
      const s = sheetNamed(ref.sheet);
      if (!s) throw new FormulaError('#REF!');
      const { maxR, maxC } = bounds.get(s.id)!;
      // Whole-column/row ranges (A:A) reach MAX_ROW; only walk the part that has data.
      const r1 = Math.min(ref.to.row - 1, maxR);
      const c1 = Math.min(ref.to.col - 1, maxC);
      const out: unknown[][] = [];
      for (let r = ref.from.row - 1; r <= r1; r++) {
        const row: unknown[] = [];
        for (let c = ref.from.col - 1; c <= c1; c++) row.push(toEngine(valueOf(s, toKey(r, c))));
        out.push(row);
      }
      return out;
    },
  });

  function valueOf(sheet: Sheet, key: string): Value {
    const id = `${sheet.id}!${key}`;
    const hit = cache.get(id);
    if (hit !== undefined) return hit;
    const raw = sheet.cells[key]?.raw ?? '';
    let v: Value;
    if (!isFormula(raw)) {
      v = parseLiteral(raw);
    } else if (visiting.has(id)) {
      return { error: '#CIRC!' };
    } else {
      visiting.add(id);
      try {
        const { r, c } = parseKey(key);
        v = fromEngine(parser.parse(translateFormula(raw.slice(1)), { sheet: sheet.name, row: r + 1, col: c + 1 }));
      } catch (e) {
        v = { error: e instanceof FormulaError ? (e as { error: string }).error : '#ERROR!' };
      } finally {
        visiting.delete(id);
      }
    }
    cache.set(id, v);
    return v;
  }

  const result: Computed = {};
  for (const sheet of wb.sheets) {
    const values: Record<string, Value> = {};
    for (const key of Object.keys(sheet.cells)) values[key] = valueOf(sheet, key);
    result[sheet.id] = values;
  }
  return result;
}
