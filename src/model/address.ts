// Rows and columns are 0-based everywhere in the app; A1 addresses are the storage keys.

export function colName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export const toKey = (r: number, c: number) => colName(c) + (r + 1);

export function parseKey(key: string): { r: number; c: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(key);
  if (!m) throw new Error(`Dirección inválida: ${key}`);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: c - 1 };
}

export interface Pos {
  r: number;
  c: number;
}

export interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export const rectOf = (a: Pos, b: Pos): Rect => ({
  r0: Math.min(a.r, b.r),
  c0: Math.min(a.c, b.c),
  r1: Math.max(a.r, b.r),
  c1: Math.max(a.c, b.c),
});

export const inRect = (rect: Rect, r: number, c: number) =>
  r >= rect.r0 && r <= rect.r1 && c >= rect.c0 && c <= rect.c1;

export function rectLabel(rect: Rect): string {
  const a = toKey(rect.r0, rect.c0);
  return rect.r0 === rect.r1 && rect.c0 === rect.c1 ? a : `${a}:${toKey(rect.r1, rect.c1)}`;
}

export function* rectKeys(rect: Rect): Generator<string> {
  for (let r = rect.r0; r <= rect.r1; r++) for (let c = rect.c0; c <= rect.c1; c++) yield toKey(r, c);
}
