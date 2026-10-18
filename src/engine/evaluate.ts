import FormulaParser from 'fast-formula-parser';
import { parseKey, toKey } from '../model/address';
import type { Sheet, Value, Workbook } from '../model/types';
import { isError } from '../model/types';
import { translateFormula } from './translate';
import { EXTRA_FUNCTIONS } from './functions';

const { FormulaError, DepParser, FormulaHelpers } = FormulaParser;

// Excel reads an empty cell as 0 wherever a number is expected (=POWER(A1,2) with A1 empty is 0);
// the engine rejects it instead, so teach it the Excel rule once.
if (!FormulaHelpers.__cellsBlankIsZero) {
  const acceptNumber = FormulaHelpers.acceptNumber;
  // A cell reference to an empty cell arrives either as null or as a { value: null } wrapper.
  const isBlank = (obj: unknown) =>
    obj === null || obj === undefined || (typeof obj === 'object' && !Array.isArray(obj) && !(obj instanceof FormulaError) && 'value' in obj && (obj as { value: unknown }).value == null);
  FormulaHelpers.acceptNumber = function (this: unknown, obj: unknown, ...rest: unknown[]) {
    return isBlank(obj) ? 0 : acceptNumber.call(this, obj, ...rest);
  };
  FormulaHelpers.__cellsBlankIsZero = true;
}

/** Computed values for every non-empty cell, by sheet id then A1 key. */
export type Computed = Record<string, Record<string, Value>>;

export interface Evaluation {
  values: Computed;
  /**
   * Cells ("sheetId!A1") whose formula Cells can't calculate yet (an unsupported function,
   * for example); they show the last value Excel or Google Sheets calculated instead.
   */
  approx: Set<string>;
}

export function parseLiteral(raw: string): Value {
  // A leading apostrophe means "this is text", like in Excel: '00123 stays 00123.
  if (raw.startsWith("'")) return raw.slice(1);
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
  else if (/^-?\d+(\.\d+)?e[+-]?\d+$/i.test(body)) n = Number(body);
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

type Ref = { sheet: string; row: number; col: number } | { sheet: string; from: { row: number; col: number }; to: { row: number; col: number } };

/** Parses a named range's target such as "'Hoja 1'!$A$2:$B$9" or "Precios!C3". */
export function parseRef(text: string): Ref | null {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(text.trim());
  if (!m) return null;
  const sheet = (m[1] ?? m[2]).replace(/''/g, "'");
  const a = parseKey(m[3].toUpperCase() + m[4]);
  if (!m[5]) return { sheet, row: a.r + 1, col: a.c + 1 };
  const b = parseKey(m[5].toUpperCase() + m[6]);
  return { sheet, from: { row: a.r + 1, col: a.c + 1 }, to: { row: b.r + 1, col: b.c + 1 } };
}

export function evaluateWorkbook(wb: Workbook): Evaluation {
  const names = new Map(Object.entries(wb.names ?? {}).map(([k, v]) => [k.toLowerCase(), v]));
  const approx = new Set<string>();
  const byName = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s]));
  const bounds = new Map(wb.sheets.map((s) => [s.id, usedBounds(s)]));
  const cache = new Map<string, Value>();
  const onStack = new Set<string>();
  const circular = new Set<string>();

  const sheetNamed = (name: string) => byName.get(String(name).toLowerCase());
  const nodeId = (sheet: Sheet, key: string) => `${sheet.id}!${key}`;
  const onVariable = (name: string) => {
    const target = names.get(name.toLowerCase());
    const ref = target ? parseRef(target) : null;
    if (!ref) throw new FormulaError('#NAME?');
    return ref;
  };

  // Formula cells per sheet, to find which ones a range like A1:C50 depends on.
  const formulaCells = new Map(
    wb.sheets.map((s) => [s.id, Object.keys(s.cells).filter((k) => isFormula(s.cells[k].raw)).map((k) => ({ key: k, ...parseKey(k) }))]),
  );

  // The engine can't safely run one formula from inside another, so each formula's inputs
  // are found first (DepParser) and calculated before it, in dependency order.
  const depParser = new DepParser({ onVariable });
  function depsOf(sheet: Sheet, key: string): [Sheet, string][] {
    const raw = sheet.cells[key].raw;
    const { r, c } = parseKey(key);
    let refs: Ref[];
    try {
      refs = depParser.parse(translateFormula(raw.slice(1)), { sheet: sheet.name, row: r + 1, col: c + 1 });
    } catch {
      return [];
    }
    const out: [Sheet, string][] = [];
    for (const ref of refs) {
      const s = sheetNamed(ref.sheet);
      if (!s) continue;
      if ('from' in ref) {
        for (const f of formulaCells.get(s.id)!) {
          if (f.r >= ref.from.row - 1 && f.r <= ref.to.row - 1 && f.c >= ref.from.col - 1 && f.c <= ref.to.col - 1) out.push([s, f.key]);
        }
      } else {
        const k = toKey(ref.row - 1, ref.col - 1);
        if (isFormula(s.cells[k]?.raw ?? '')) out.push([s, k]);
      }
    }
    return out;
  }

  const parser = new FormulaParser({
    functions: EXTRA_FUNCTIONS,
    onVariable,
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

  function evalOne(sheet: Sheet, key: string): Value {
    const cell = sheet.cells[key];
    const id = nodeId(sheet, key);
    // References to other files ("[1]Hoja!A1") can't be followed: use Excel's last value.
    if (/\[\d+\]/.test(cell.raw) && cell.cached !== undefined) {
      approx.add(id);
      return cell.cached;
    }
    let v: Value;
    let unsupported = false;
    try {
      const { r, c } = parseKey(key);
      v = fromEngine(parser.parse(translateFormula(cell.raw.slice(1)), { sheet: sheet.name, row: r + 1, col: c + 1 }));
    } catch (e) {
      const message = String((e as { details?: { message?: string } })?.details?.message ?? '');
      // An unknown function is #NAME?, as in Excel; anything else the engine can't read is #ERROR!.
      v = { error: /is not implemented/.test(message) ? '#NAME?' : e instanceof FormulaError ? (e as { error: string }).error : '#ERROR!' };
      unsupported = v.error === '#NAME?' || v.error === '#ERROR!';
    }
    // A formula Cells can't run yet (not a real spreadsheet error like #DIV/0!): if the file
    // came with Excel's own result, show that instead.
    if (unsupported && cell.cached !== undefined) {
      approx.add(id);
      return cell.cached;
    }
    return v;
  }

  /** Calculates a formula after everything it depends on, without recursion. */
  function evaluateTree(sheet: Sheet, key: string) {
    type Frame = { sheet: Sheet; key: string; id: string; deps: [Sheet, string][]; i: number };
    const stack: Frame[] = [];
    const push = (s: Sheet, k: string) => {
      const id = nodeId(s, k);
      onStack.add(id);
      stack.push({ sheet: s, key: k, id, deps: depsOf(s, k), i: 0 });
    };
    push(sheet, key);
    while (stack.length) {
      const top = stack[stack.length - 1];
      if (top.i < top.deps.length) {
        const [ds, dk] = top.deps[top.i++];
        const did = nodeId(ds, dk);
        if (cache.has(did)) continue;
        if (onStack.has(did)) {
          const from = stack.findIndex((f) => f.id === did);
          for (let j = Math.max(0, from); j < stack.length; j++) circular.add(stack[j].id);
          continue;
        }
        push(ds, dk);
        continue;
      }
      stack.pop();
      onStack.delete(top.id);
      cache.set(top.id, circular.has(top.id) ? { error: '#CIRC!' } : evalOne(top.sheet, top.key));
    }
  }

  function valueOf(sheet: Sheet, key: string): Value {
    const id = nodeId(sheet, key);
    const hit = cache.get(id);
    if (hit !== undefined) return hit;
    const raw = sheet.cells[key]?.raw ?? '';
    if (!isFormula(raw)) {
      const v = parseLiteral(raw);
      cache.set(id, v);
      return v;
    }
    // Only reached for references the dependency pass couldn't see in advance.
    if (onStack.has(id)) return { error: '#CIRC!' };
    evaluateTree(sheet, key);
    return cache.get(id)!;
  }

  const result: Computed = {};
  for (const sheet of wb.sheets) {
    const values: Record<string, Value> = {};
    for (const key of Object.keys(sheet.cells)) values[key] = valueOf(sheet, key);
    result[sheet.id] = values;
  }
  return { values: result, approx };
}
