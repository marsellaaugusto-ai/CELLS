export type NumFmt = 'general' | 'number2' | 'currency' | 'percent';
export type Align = 'left' | 'center' | 'right' | 'justify' | 'distributed' | 'fill' | 'centerContinuous';
export type Tint = 'none' | 'blue' | 'green' | 'teal' | 'orange' | 'yellow' | 'red' | 'purple';

export type VAlign = 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
/** Excel's border styles, kept by name so they go back unchanged. */
export type BorderStyle =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';

export interface BorderSide {
  style: BorderStyle;
  /** '#rrggbb' */
  color?: string;
}

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  align?: Align;
  valign?: VAlign;
  wrap?: boolean;
  /** Presets offered in the Formato panel. */
  numFmt?: NumFmt;
  /** Excel number format code, e.g. 'dd/mm/yyyy' or '#,##0.00 "kg"'. Wins over numFmt. */
  numFmtCode?: string;
  /** Preset tile colors offered in the Formato panel. */
  tint?: Tint;
  /** Exact background color ('#rrggbb'), e.g. from an imported file. Wins over tint. */
  fill?: string;
  /** Exact text color ('#rrggbb'). */
  color?: string;
  /** Points, like Excel. */
  fontSize?: number;
  fontFamily?: string;
  border?: { top?: BorderSide; right?: BorderSide; bottom?: BorderSide; left?: BorderSide };
}

export interface Validation {
  /** Allowed values when the rule is a plain list, offered as suggestions while typing. */
  list: string[] | null;
  /** The rule exactly as it came from Excel, written back unchanged on export. */
  excel?: unknown;
}

/** A piece of text with its own format, for cells that mix formats (m², bold words…). */
export interface RichRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** '#rrggbb' */
  color?: string;
  /** Points. */
  size?: number;
  font?: string;
  vertAlign?: 'superscript' | 'subscript';
}

export interface Cell {
  /** What the person typed: a value, or a formula starting with '='. A leading ' forces text. */
  raw: string;
  fmt?: CellFormat;
  /** Comment shown on hover (Excel "note"). */
  note?: string;
  link?: string;
  /** Last value Excel or Google Sheets calculated, used when Cells can't calculate the formula itself. */
  cached?: number | string | boolean | ErrorValue;
  validation?: Validation;
  /** Mixed formats inside the text, as Excel stores them. Dropped when the text is edited. */
  rich?: RichRun[];
}

export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface SheetLayout {
  /** Pixel widths/heights by 0-based column/row index; missing = default. */
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
  hiddenCols?: number[];
  hiddenRows?: number[];
  merges?: Range[];
  frozen?: { rows: number; cols: number };
  /** Conditional formatting rules from Excel, kept so they survive a round trip. */
  conditional?: unknown[];
  /** Excel page setup (paper, orientation, margins, scale), kept for the round trip. */
  pageSetup?: Record<string, unknown>;
}

export interface Sheet {
  id: string;
  name: string;
  /** Keyed by A1-style address, e.g. "B3". Missing key = empty cell. */
  cells: Record<string, Cell>;
  rows: number;
  cols: number;
  layout?: SheetLayout;
}

export interface Workbook {
  id: string;
  name: string;
  sheets: Sheet[];
  updatedAt: string;
  /** Named ranges: name → reference, e.g. { Precios: "'Hoja 1'!$D$2:$D$9" }. */
  names?: Record<string, string>;
  /** The file's base font (Excel's "Normal" style), used by cells that don't set their own. */
  defaultFont?: { name: string; size: number };
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
