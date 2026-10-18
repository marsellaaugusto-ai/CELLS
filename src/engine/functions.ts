import FormulaParser from 'fast-formula-parser';

// Common Excel / Google Sheets functions the engine lacks, written to match their documented
// behavior. Arguments arrive wrapped by the parser: { value, isArray, isRangeRef, isCellRef }.

const { FormulaError } = FormulaParser;

interface Arg {
  value: unknown;
  isRangeRef?: boolean;
  isArray?: boolean;
  omitted?: boolean;
}

const isErr = (v: unknown) => v instanceof FormulaError;

function flat(arg: Arg | undefined): unknown[] {
  if (!arg) return [];
  const v = arg.value;
  return Array.isArray(v) ? (v as unknown[][]).flat() : [v];
}

function grid(arg: Arg): unknown[][] {
  const v = arg.value;
  return Array.isArray(v) ? (v as unknown[][]) : [[v]];
}

function num(arg: Arg | undefined, fallback?: number): number {
  if (!arg || arg.omitted || arg.value === null || arg.value === undefined) {
    if (fallback !== undefined) return fallback;
    throw new FormulaError('#VALUE!');
  }
  const v = Array.isArray(arg.value) ? (arg.value as unknown[][])[0][0] : arg.value;
  if (isErr(v)) throw v;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(String(v).replace(',', '.'));
  if (String(v).trim() === '' || isNaN(n)) throw new FormulaError('#VALUE!');
  return n;
}

const scalar = (arg: Arg | undefined) => (arg ? (Array.isArray(arg.value) ? (arg.value as unknown[][])[0][0] : arg.value) : undefined);

const numbers = (args: Arg[]) => args.flatMap(flat).filter((v): v is number => typeof v === 'number');

/** Excel criteria: 5, ">5", "<>x", "=", "Ac*", "?ero"; text comparisons ignore case. */
export function matchesCriteria(value: unknown, criteria: unknown): boolean {
  if (typeof criteria === 'number' || typeof criteria === 'boolean') return value === criteria;
  const c = String(criteria ?? '');
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const rhsNum = rhs.trim() !== '' && !isNaN(Number(rhs)) ? Number(rhs) : null;
  const empty = value === null || value === undefined || value === '';
  if (rhs === '') return op === '=' ? empty : op === '<>' ? !empty : false;
  if (rhsNum !== null && typeof value === 'number') {
    switch (op) {
      case '=': return value === rhsNum;
      case '<>': return value !== rhsNum;
      case '<': return value < rhsNum;
      case '>': return value > rhsNum;
      case '<=': return value <= rhsNum;
      case '>=': return value >= rhsNum;
    }
  }
  if (op === '=' || op === '<>') {
    const pattern = new RegExp(
      '^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0000').replace(/~\?/g, '\u0001').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0000/g, '\\*').replace(/\u0001/g, '\\?') + '$',
      'is',
    );
    const hit = !empty && typeof value !== 'number' ? pattern.test(String(value)) : false;
    return op === '=' ? hit : !hit;
  }
  if (typeof value !== 'string') return false;
  const cmp = value.localeCompare(rhs, undefined, { sensitivity: 'base' });
  return op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0;
}

function lookupIndex(needle: unknown, haystack: unknown[], matchMode: number, searchMode: number): number {
  const order = searchMode === -1 ? [...haystack.keys()].reverse() : [...haystack.keys()];
  const eq = (a: unknown, b: unknown) =>
    typeof a === 'string' && typeof b === 'string' ? a.localeCompare(b, undefined, { sensitivity: 'base' }) === 0 : a === b;
  if (matchMode === 2) {
    for (const i of order) if (matchesCriteria(haystack[i], `=${String(needle)}`)) return i;
    return -1;
  }
  for (const i of order) if (eq(haystack[i], needle)) return i;
  if (matchMode === 0) return -1;
  // -1: exact or next smaller; 1: exact or next larger
  let best = -1;
  for (const i of order) {
    const v = haystack[i];
    if (typeof v !== typeof needle || v === null) continue;
    const below = (v as number) < (needle as number);
    if (matchMode === -1 && below && (best === -1 || (v as number) > (haystack[best] as number))) best = i;
    if (matchMode === 1 && !below && (best === -1 || (v as number) < (haystack[best] as number))) best = i;
  }
  return best;
}

function countIfs(args: Arg[]): boolean[] {
  if (args.length % 2) throw new FormulaError('#VALUE!');
  let mask: boolean[] | null = null;
  for (let i = 0; i < args.length; i += 2) {
    const vals = flat(args[i]);
    const crit = scalar(args[i + 1]);
    if (mask && mask.length !== vals.length) throw new FormulaError('#VALUE!');
    mask = vals.map((v, j) => (mask ? mask[j] : true) && matchesCriteria(v, crit));
  }
  return mask ?? [];
}

function stdev(xs: number[], sample: boolean) {
  if (xs.length < (sample ? 2 : 1)) throw new FormulaError('#DIV/0!');
  const mean = xs.reduce((a, b) => a + b, 0) / xs.length;
  const ss = xs.reduce((a, b) => a + (b - mean) ** 2, 0);
  return ss / (xs.length - (sample ? 1 : 0));
}

function pmt(rate: number, n: number, pv: number, fv = 0, type = 0) {
  if (rate === 0) return -(pv + fv) / n;
  const f = (1 + rate) ** n;
  return (-(pv * f + fv) * rate) / ((1 + rate * type) * (f - 1));
}

export const EXTRA_FUNCTIONS: Record<string, (...args: Arg[]) => unknown> = {
  COUNTA: (...args) => args.flatMap(flat).filter((v) => v !== null && v !== undefined && v !== '').length,
  COUNTBLANK: (...args) => args.flatMap(flat).filter((v) => v === null || v === undefined || v === '').length,
  COUNTIFS: (...args) => countIfs(args).filter(Boolean).length,
  UPPER: (t) => String(scalar(t) ?? '').toUpperCase(),
  VALUE: (t) => num(t),
  'STDEV': (...args) => Math.sqrt(stdev(numbers(args), true)),
  'VAR': (...args) => stdev(numbers(args), true),
  MODE: (...args) => {
    const counts = new Map<number, number>();
    let best: number | null = null;
    for (const x of numbers(args)) {
      const c = (counts.get(x) ?? 0) + 1;
      counts.set(x, c);
      if (c > 1 && (best === null || c > counts.get(best)!)) best = x;
    }
    if (best === null) throw new FormulaError('#N/A');
    return best;
  },
  RANK: (x, ref, order) => {
    const v = num(x);
    const xs = numbers([ref]);
    if (!xs.includes(v)) throw new FormulaError('#N/A');
    const asc = num(order, 0) !== 0;
    return 1 + xs.filter((y) => (asc ? y < v : y > v)).length;
  },
  XMATCH: (needle, list, match, search) => {
    const i = lookupIndex(scalar(needle), flat(list), num(match, 0), num(search, 1));
    if (i < 0) throw new FormulaError('#N/A');
    return i + 1;
  },
  XLOOKUP: (needle, lookup, ret, notFound, match, search) => {
    const i = lookupIndex(scalar(needle), flat(lookup), num(match, 0), num(search, 1));
    if (i < 0) {
      if (notFound && !notFound.omitted) return scalar(notFound);
      throw new FormulaError('#N/A');
    }
    const g = grid(ret);
    // Vertical lookup range → same row of ret; horizontal → same column.
    const vertical = grid(lookup).length > 1;
    const v = vertical ? g[i]?.[0] : g[0]?.[i];
    return v ?? 0;
  },
  PMT: (rate, n, pv, fv, type) => pmt(num(rate), num(n), num(pv), num(fv, 0), num(type, 0)),
  FV: (rate, n, payment, pv, type) => {
    const r = num(rate), k = num(n), p = num(payment), v = num(pv, 0), t = num(type, 0);
    if (r === 0) return -(v + p * k);
    const f = (1 + r) ** k;
    return -(v * f + (p * (1 + r * t) * (f - 1)) / r);
  },
  PV: (rate, n, payment, fv, type) => {
    const r = num(rate), k = num(n), p = num(payment), v = num(fv, 0), t = num(type, 0);
    if (r === 0) return -(v + p * k);
    const f = (1 + r) ** k;
    return -(v + (p * (1 + r * t) * (f - 1)) / r) / f;
  },
  NPV: (rate, ...flows) => {
    const r = num(rate);
    return numbers(flows).reduce((acc, cf, i) => acc + cf / (1 + r) ** (i + 1), 0);
  },
};
