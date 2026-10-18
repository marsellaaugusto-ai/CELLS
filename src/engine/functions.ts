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

/**
 * Numbers for MAX, MIN, AVERAGE-style functions, with Excel's rules: inside ranges only real
 * numbers count (text, booleans and blanks are skipped); typed-in arguments are converted
 * (TRUE → 1, "3" → 3); errors stop the calculation.
 */
function nums(args: Arg[], { textAsZero = false } = {}): number[] {
  const out: number[] = [];
  for (const a of args) {
    if (!a || a.omitted) continue;
    const fromRef = a.isRangeRef || a.isArray || (a as { isCellRef?: boolean }).isCellRef;
    for (const v of flat(a)) {
      if (isErr(v)) throw v;
      if (typeof v === 'number') out.push(v);
      else if (fromRef) {
        if (textAsZero && typeof v === 'string') out.push(0);
        else if (textAsZero && typeof v === 'boolean') out.push(v ? 1 : 0);
      } else if (typeof v === 'boolean') out.push(v ? 1 : 0);
      else if (typeof v === 'string') {
        const n = Number(v.replace(',', '.'));
        if (v.trim() === '' || isNaN(n)) throw new FormulaError('#VALUE!');
        out.push(n);
      }
    }
  }
  return out;
}

const sorted = (xs: number[]) => [...xs].sort((a, b) => a - b);

function kth(xs: number[], k: number, largest: boolean) {
  const s = sorted(xs);
  if (k < 1 || k > s.length || !Number.isFinite(k)) throw new FormulaError('#NUM!');
  return largest ? s[s.length - Math.ceil(k)] : s[Math.ceil(k) - 1];
}

function percentileInc(xs: number[], p: number) {
  if (!xs.length || p < 0 || p > 1) throw new FormulaError('#NUM!');
  const s = sorted(xs);
  const h = (s.length - 1) * p;
  const lo = Math.floor(h);
  return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]);
}

function percentileExc(xs: number[], p: number) {
  const s = sorted(xs);
  const h = (s.length + 1) * p - 1;
  if (!s.length || h < 0 || h > s.length - 1) throw new FormulaError('#NUM!');
  const lo = Math.floor(h);
  return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]);
}

function mean(xs: number[]) {
  if (!xs.length) throw new FormulaError('#DIV/0!');
  return xs.reduce((a, b) => a + b, 0) / xs.length;
}

/** Values of the cells in the first range where every criteria pair matches (…IFS functions). */
function ifsPick(target: Arg, pairs: Arg[]): unknown[] {
  const vals = flat(target);
  const mask = countIfs(pairs);
  if (mask.length !== vals.length) throw new FormulaError('#VALUE!');
  return vals.filter((_, i) => mask[i]);
}

const ifsNumbers = (target: Arg, pairs: Arg[]) => ifsPick(target, pairs).filter((v): v is number => typeof v === 'number');

function pairStats(ys: Arg, xs: Arg) {
  const a = flat(ys);
  const b = flat(xs);
  if (a.length !== b.length) throw new FormulaError('#N/A');
  const pts = a.map((y, i) => [b[i], y]).filter((p): p is [number, number] => typeof p[0] === 'number' && typeof p[1] === 'number');
  if (pts.length < 2) throw new FormulaError('#DIV/0!');
  const mx = mean(pts.map((p) => p[0]));
  const my = mean(pts.map((p) => p[1]));
  let sxy = 0, sxx = 0, syy = 0;
  for (const [x, y] of pts) {
    sxy += (x - mx) * (y - my);
    sxx += (x - mx) ** 2;
    syy += (y - my) ** 2;
  }
  return { sxy, sxx, syy, mx, my };
}

/** Position for MATCH: 0 exact (wildcards allowed), 1 largest ≤ value, -1 smallest ≥ value. */
function matchIndex(needle: unknown, list: unknown[], type: number): number {
  if (type === 0) {
    if (typeof needle === 'string') return list.findIndex((v) => matchesCriteria(v, `=${needle}`));
    return list.findIndex((v) => v === needle);
  }
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (typeof v !== typeof needle || v === null) continue;
    const cmp = typeof v === 'string' ? v.localeCompare(needle as string, undefined, { sensitivity: 'base' }) : (v as number) - (needle as number);
    if (type > 0 ? cmp <= 0 : cmp >= 0) best = i;
    else break; // the list is expected sorted, as in Excel
  }
  return best;
}

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

/** Logical values for AND/OR: ranges contribute booleans and numbers, text is skipped, errors stop. */
function logicals(args: Arg[]): boolean[] {
  const out: boolean[] = [];
  for (const a of args) {
    if (!a || a.omitted) continue;
    const fromRef = a.isRangeRef || a.isArray || (a as { isCellRef?: boolean }).isCellRef;
    for (const v of flat(a)) {
      if (isErr(v)) throw v;
      if (typeof v === 'boolean') out.push(v);
      else if (typeof v === 'number') out.push(v !== 0);
      else if (typeof v === 'string' && !fromRef) {
        const u = v.toUpperCase();
        if (u === 'TRUE' || u === 'VERDADERO') out.push(true);
        else if (u === 'FALSE' || u === 'FALSO') out.push(false);
        else throw new FormulaError('#VALUE!');
      }
    }
  }
  if (!out.length) throw new FormulaError('#VALUE!');
  return out;
}

export const EXTRA_FUNCTIONS: Record<string, (...args: Arg[]) => unknown> = {
  // The engine's versions of these skip or garble errors inside ranges; Excel passes the error
  // on (if one cell of a SUM range is #DIV/0!, the SUM is #DIV/0!).
  SUM: (...args) => nums(args).reduce((a, b) => a + b, 0),
  AVERAGE: (...args) => mean(nums(args)),
  AVERAGEA: (...args) => mean(nums(args, { textAsZero: true })),
  PRODUCT: (...args) => {
    const xs = nums(args);
    return xs.length ? xs.reduce((a, b) => a * b, 1) : 0;
  },
  SUMSQ: (...args) => nums(args).reduce((a, b) => a + b * b, 0),
  SUMPRODUCT: (...arrays) => {
    const lists = arrays.map((a) => flat(a));
    const n = lists[0]?.length ?? 0;
    if (lists.some((l) => l.length !== n)) throw new FormulaError('#VALUE!');
    let total = 0;
    for (let i = 0; i < n; i++) {
      let prod = 1;
      for (const l of lists) {
        const v = l[i];
        if (isErr(v)) throw v;
        prod *= typeof v === 'number' ? v : 0;
      }
      total += prod;
    }
    return total;
  },
  AND: (...args) => logicals(args).every(Boolean),
  OR: (...args) => logicals(args).some(Boolean),
  // The engine declares these but leaves them empty; written here following Excel's documentation.
  MAX: (...args) => {
    const xs = nums(args);
    return xs.length ? Math.max(...xs) : 0;
  },
  MIN: (...args) => {
    const xs = nums(args);
    return xs.length ? Math.min(...xs) : 0;
  },
  MAXA: (...args) => {
    const xs = nums(args, { textAsZero: true });
    return xs.length ? Math.max(...xs) : 0;
  },
  MINA: (...args) => {
    const xs = nums(args, { textAsZero: true });
    return xs.length ? Math.min(...xs) : 0;
  },
  MEDIAN: (...args) => {
    const s = sorted(nums(args));
    if (!s.length) throw new FormulaError('#NUM!');
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  },
  LARGE: (range, k) => kth(nums([range]), num(k), true),
  SMALL: (range, k) => kth(nums([range]), num(k), false),
  'PERCENTILE.INC': (range, p) => percentileInc(nums([range]), num(p)),
  'PERCENTILE.EXC': (range, p) => percentileExc(nums([range]), num(p)),
  PERCENTILE: (range, p) => percentileInc(nums([range]), num(p)),
  'QUARTILE.INC': (range, q) => percentileInc(nums([range]), Math.trunc(num(q)) / 4),
  'QUARTILE.EXC': (range, q) => percentileExc(nums([range]), Math.trunc(num(q)) / 4),
  QUARTILE: (range, q) => percentileInc(nums([range]), Math.trunc(num(q)) / 4),
  'STDEV.S': (...args) => Math.sqrt(stdev(nums(args), true)),
  'STDEV.P': (...args) => Math.sqrt(stdev(nums(args), false)),
  STDEVP: (...args) => Math.sqrt(stdev(nums(args), false)),
  'VAR.S': (...args) => stdev(nums(args), true),
  'VAR.P': (...args) => stdev(nums(args), false),
  VARP: (...args) => stdev(nums(args), false),
  TRIMMEAN: (range, pct) => {
    const s = sorted(nums([range]));
    const p = num(pct);
    if (p < 0 || p >= 1) throw new FormulaError('#NUM!');
    const cut = Math.floor((s.length * p) / 2);
    return mean(s.slice(cut, s.length - cut));
  },
  SKEW: (...args) => {
    const xs = nums(args);
    const n = xs.length;
    if (n < 3) throw new FormulaError('#DIV/0!');
    const m = mean(xs);
    const sd = Math.sqrt(stdev(xs, true));
    return (n / ((n - 1) * (n - 2))) * xs.reduce((a, x) => a + ((x - m) / sd) ** 3, 0);
  },
  PERMUT: (n, k) => {
    const a = Math.trunc(num(n)), b = Math.trunc(num(k));
    if (a < 0 || b < 0 || a < b) throw new FormulaError('#NUM!');
    let r = 1;
    for (let i = 0; i < b; i++) r *= a - i;
    return r;
  },
  SLOPE: (ys, xs) => {
    const s = pairStats(ys, xs);
    if (!s.sxx) throw new FormulaError('#DIV/0!');
    return s.sxy / s.sxx;
  },
  PEARSON: (a, b) => {
    const s = pairStats(a, b);
    return s.sxy / Math.sqrt(s.sxx * s.syy);
  },
  RSQ: (ys, xs) => {
    const s = pairStats(ys, xs);
    return (s.sxy * s.sxy) / (s.sxx * s.syy);
  },
  'RANK.EQ': (x, ref, order) => {
    const v = num(x);
    const xs = nums([ref]);
    if (!xs.includes(v)) throw new FormulaError('#N/A');
    const asc = num(order, 0) !== 0;
    return 1 + xs.filter((y) => (asc ? y < v : y > v)).length;
  },
  'RANK.AVG': (x, ref, order) => {
    const v = num(x);
    const xs = nums([ref]);
    const ties = xs.filter((y) => y === v).length;
    if (!ties) throw new FormulaError('#N/A');
    const asc = num(order, 0) !== 0;
    return 1 + xs.filter((y) => (asc ? y < v : y > v)).length + (ties - 1) / 2;
  },
  'MODE.SNGL': (...args) => EXTRA_FUNCTIONS.MODE(...args),
  SUMIFS: (target, ...pairs) => ifsNumbers(target, pairs).reduce((a, b) => a + b, 0),
  AVERAGEIFS: (target, ...pairs) => mean(ifsNumbers(target, pairs)),
  MAXIFS: (target, ...pairs) => {
    const xs = ifsNumbers(target, pairs);
    return xs.length ? Math.max(...xs) : 0;
  },
  MINIFS: (target, ...pairs) => {
    const xs = ifsNumbers(target, pairs);
    return xs.length ? Math.min(...xs) : 0;
  },
  MATCH: (needle, list, type) => {
    const i = matchIndex(scalar(needle), flat(list), num(type, 1));
    if (i < 0) throw new FormulaError('#N/A');
    return i + 1;
  },
  // The engine passes its context first to CHOOSE (it's on its list of context functions).
  CHOOSE: (_context, index, ...options) => {
    const i = Math.trunc(num(index));
    if (i < 1 || i > options.length) throw new FormulaError('#VALUE!');
    return scalar(options[i - 1]) ?? 0;
  },
  SWITCH: (expr, ...rest) => {
    const v = scalar(expr);
    for (let i = 0; i + 1 < rest.length; i += 2) {
      const c = scalar(rest[i]);
      if (c === v || (typeof c === 'string' && typeof v === 'string' && c.toLowerCase() === v.toLowerCase())) return scalar(rest[i + 1]);
    }
    if (rest.length % 2) return scalar(rest[rest.length - 1]);
    throw new FormulaError('#N/A');
  },
  SUBSTITUTE: (text, oldText, newText, instance) => {
    const t = String(scalar(text) ?? '');
    const o = String(scalar(oldText) ?? '');
    const n = String(scalar(newText) ?? '');
    if (!o) return t;
    if (!instance || instance.omitted) return t.split(o).join(n);
    const which = Math.trunc(num(instance));
    let idx = -1;
    for (let k = 0; k < which; k++) {
      idx = t.indexOf(o, idx + 1);
      if (idx < 0) return t;
    }
    return t.slice(0, idx) + n + t.slice(idx + o.length);
  },
  TEXTJOIN: (delim, ignoreEmpty, ...parts) => {
    const skip = Boolean(scalar(ignoreEmpty));
    const vals = parts.flatMap(flat).map((v) => (v === null || v === undefined ? '' : String(v)));
    return (skip ? vals.filter((v) => v !== '') : vals).join(String(scalar(delim) ?? ''));
  },
  SUBTOTAL: (fn, ...refs) => {
    const code = Math.trunc(num(fn)) % 100;
    const xs = nums(refs);
    switch (code) {
      case 1: return mean(xs);
      case 2: return xs.length;
      case 3: return refs.flatMap(flat).filter((v) => v !== null && v !== undefined && v !== '').length;
      case 4: return xs.length ? Math.max(...xs) : 0;
      case 5: return xs.length ? Math.min(...xs) : 0;
      case 6: return xs.reduce((a, b) => a * b, 1);
      case 7: return Math.sqrt(stdev(xs, true));
      case 8: return Math.sqrt(stdev(xs, false));
      case 9: return xs.reduce((a, b) => a + b, 0);
      case 10: return stdev(xs, true);
      case 11: return stdev(xs, false);
      default: throw new FormulaError('#VALUE!');
    }
  },
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
