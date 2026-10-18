// Lets people write formulas the way Spanish Excel / Google Sheets does:
// Spanish function names (SUMA, PROMEDIO, SI…) and ';' as argument separator.
// The engine itself only understands the English names and ','.

export const SPANISH_FUNCTIONS: Record<string, string> = {
  SUMA: 'SUM',
  PROMEDIO: 'AVERAGE',
  CONTAR: 'COUNT',
  'CONTAR.SI': 'COUNTIF',
  'SUMAR.SI': 'SUMIF',
  'PROMEDIO.SI': 'AVERAGEIF',
  SI: 'IF',
  'SI.ERROR': 'IFERROR',
  'SI.CONJUNTO': 'IFS',
  Y: 'AND',
  O: 'OR',
  NO: 'NOT',
  BUSCARV: 'VLOOKUP',
  BUSCARH: 'HLOOKUP',
  INDICE: 'INDEX',
  REDONDEAR: 'ROUND',
  'REDONDEAR.MAS': 'ROUNDUP',
  'REDONDEAR.MENOS': 'ROUNDDOWN',
  ENTERO: 'INT',
  RESIDUO: 'MOD',
  POTENCIA: 'POWER',
  RAIZ: 'SQRT',
  CONCATENAR: 'CONCATENATE',
  IZQUIERDA: 'LEFT',
  DERECHA: 'RIGHT',
  EXTRAE: 'MID',
  LARGO: 'LEN',
  MAYUSC: 'UPPER',
  MINUSC: 'LOWER',
  ESPACIOS: 'TRIM',
  TEXTO: 'TEXT',
  HOY: 'TODAY',
  AHORA: 'NOW',
  FECHA: 'DATE',
  'AÑO': 'YEAR',
  MES: 'MONTH',
  DIA: 'DAY',
  ALEATORIO: 'RAND',
  'ALEATORIO.ENTRE': 'RANDBETWEEN',
  ESBLANCO: 'ISBLANK',
  ESNUMERO: 'ISNUMBER',
  ESTEXTO: 'ISTEXT',
  PRODUCTO: 'PRODUCT',
  ABS: 'ABS',
  MAX: 'MAX',
  MIN: 'MIN',
};

const IDENT_START = /[A-Za-zÁÉÍÓÚÜÑáéíóúüñ_]/;
const IDENT_PART = /[A-Za-z0-9ÁÉÍÓÚÜÑáéíóúüñ_.]/;

export function translateFormula(src: string): string {
  let out = '';
  let i = 0;
  let braces = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"' || ch === "'") {
      // Copy string literals / quoted sheet names verbatim ("" and '' are escapes).
      let j = i + 1;
      while (j < src.length) {
        if (src[j] === ch) {
          if (src[j + 1] === ch) j += 2;
          else break;
        } else j++;
      }
      out += src.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '{') braces++;
    if (ch === '}') braces--;
    if (ch === ';' && braces === 0) {
      out += ',';
      i++;
      continue;
    }
    if (IDENT_START.test(ch)) {
      let j = i + 1;
      while (j < src.length && IDENT_PART.test(src[j])) j++;
      const word = src.slice(i, j);
      let k = j;
      while (src[k] === ' ') k++;
      const english = src[k] === '(' ? SPANISH_FUNCTIONS[word.toUpperCase()] : undefined;
      out += english ?? word;
      i = j;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}
