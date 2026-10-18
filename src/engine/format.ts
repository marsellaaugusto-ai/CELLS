import SSF from 'ssf';
import type { CellFormat, Value } from '../model/types';
import { isError } from '../model/types';

const LOCALE = 'es-AR';

const general = new Intl.NumberFormat(LOCALE, { maximumFractionDigits: 10 });
const fixed2 = new Intl.NumberFormat(LOCALE, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const percent = new Intl.NumberFormat(LOCALE, { style: 'percent', maximumFractionDigits: 2 });

export function formatValue(v: Value, fmt?: CellFormat): string {
  if (v === null) return '';
  if (isError(v)) return v.error;
  if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
  if (typeof v === 'string') return v;
  if (fmt?.numFmtCode) return formatWithCode(v, fmt.numFmtCode);
  switch (fmt?.numFmt) {
    case 'number2':
      return fixed2.format(v);
    case 'currency':
      return '$ ' + fixed2.format(v);
    case 'percent':
      return percent.format(v);
    default:
      return general.format(v);
  }
}

/** Numbers line up on the right unless the cell says otherwise, like every spreadsheet. */
export function alignFor(v: Value, fmt?: CellFormat) {
  if (fmt?.align) return fmt.align;
  return typeof v === 'number' || typeof v === 'boolean' ? 'right' : 'left';
}

/**
 * Renders a number with an Excel format code ('#,##0.00', 'dd/mm/yyyy', '0%'…), then swaps
 * the separators to the Spanish style, as Excel does with a Spanish regional setting.
 * Format codes are stored in the international form, so files keep working everywhere.
 */
export function formatWithCode(v: number, code: string): string {
  let out: string;
  try {
    out = SSF.format(code, v);
  } catch {
    return general.format(v);
  }
  out = out.replace(/\d[\d,.]*\d/g, (num) => num.replace(/[.,]/g, (ch) => (ch === '.' ? ',' : '.')));
  if (/m{3}|d{3}/i.test(code)) out = out.replace(/\b[A-Z][a-z]+\b/g, (w) => SPANISH_DATE_WORDS[w] ?? w);
  return out;
}

const SPANISH_DATE_WORDS: Record<string, string> = Object.fromEntries(
  [
    ['January', 'enero'], ['February', 'febrero'], ['March', 'marzo'], ['April', 'abril'], ['May', 'mayo'], ['June', 'junio'],
    ['July', 'julio'], ['August', 'agosto'], ['September', 'septiembre'], ['October', 'octubre'], ['November', 'noviembre'],
    ['December', 'diciembre'], ['Jan', 'ene'], ['Feb', 'feb'], ['Mar', 'mar'], ['Apr', 'abr'], ['Jun', 'jun'], ['Jul', 'jul'],
    ['Aug', 'ago'], ['Sep', 'sep'], ['Oct', 'oct'], ['Nov', 'nov'], ['Dec', 'dic'], ['Monday', 'lunes'], ['Tuesday', 'martes'],
    ['Wednesday', 'miércoles'], ['Thursday', 'jueves'], ['Friday', 'viernes'], ['Saturday', 'sábado'], ['Sunday', 'domingo'],
    ['Mon', 'lun'], ['Tue', 'mar'], ['Wed', 'mié'], ['Thu', 'jue'], ['Fri', 'vie'], ['Sat', 'sáb'], ['Sun', 'dom'],
  ],
);

const CODE_COLORS: Record<string, string> = {
  red: '#ff8f87', blue: '#7cb8f5', green: '#6cc070', yellow: '#e9b326', magenta: '#e07ad8', cyan: '#5fd3e3', white: '#ffffff',
};

/** Text color a format code asks for, like the [Red] in '#,##0;[Red]-#,##0'. Tuned to read on dark tiles. */
export function codeColor(v: Value, code?: string): string | undefined {
  if (!code || typeof v !== 'number') return undefined;
  const sections = code.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  const section = v < 0 && sections[1] !== undefined ? sections[1] : v === 0 && sections[2] !== undefined ? sections[2] : sections[0];
  const m = /\[(red|blue|green|yellow|magenta|cyan|white|black)\]/i.exec(section);
  return m ? CODE_COLORS[m[1].toLowerCase()] : undefined;
}
