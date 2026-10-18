export type NumFmt = 'general' | 'number2' | 'currency' | 'percent';
export type Align = 'left' | 'center' | 'right';
export type Tint = 'none' | 'blue' | 'green' | 'teal' | 'orange' | 'yellow' | 'red' | 'purple';

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  align?: Align;
  numFmt?: NumFmt;
  tint?: Tint;
}

export interface Cell {
  raw: string;
  fmt?: CellFormat;
}

export interface Sheet {
  id: string;
  name: string;
  /** Keyed by A1-style address, e.g. "B3". Missing key = empty cell. */
  cells: Record<string, Cell>;
  rows: number;
  cols: number;
}

export interface Workbook {
  id: string;
  name: string;
  sheets: Sheet[];
  updatedAt: string;
}

export interface WorkbookSummary {
  id: string;
  name: string;
  updatedAt: string;
}

export type ErrorValue = { error: string };
export type Value = number | string | boolean | null | ErrorValue;

export const isError = (v: Value): v is ErrorValue =>
  typeof v === 'object' && v !== null && 'error' in v;

export const newId = () =>
  typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2) + Date.now().toString(36);

export function newSheet(name: string): Sheet {
  return { id: newId(), name, cells: {}, rows: 100, cols: 26 };
}

export function newWorkbook(name = 'Libro sin título'): Workbook {
  return { id: newId(), name, sheets: [newSheet('Hoja 1')], updatedAt: new Date().toISOString() };
}
