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
