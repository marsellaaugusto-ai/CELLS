import { describe, expect, it } from 'vitest';
import type { Workbook } from '../model/types';
import { evaluateWorkbook, parseLiteral } from './evaluate';
import { formatValue } from './format';
import { translateFormula } from './translate';

function book(...sheets: [string, Record<string, string>][]): Workbook {
  return {
    id: 'wb',
    name: 'Test',
    updatedAt: '',
    sheets: sheets.map(([name, raw], i) => ({
      id: `s${i}`,
      name,
      rows: 100,
      cols: 26,
      cells: Object.fromEntries(Object.entries(raw).map(([k, v]) => [k, { raw: v }])),
    })),
  };
}

describe('parseLiteral', () => {
  it('reads numbers in both decimal styles', () => {
    expect(parseLiteral('42')).toBe(42);
    expect(parseLiteral('3,5')).toBe(3.5);
    expect(parseLiteral('3.5')).toBe(3.5);
    expect(parseLiteral('1.540,25')).toBe(1540.25);
    expect(parseLiteral('15%')).toBe(0.15);
    expect(parseLiteral('VERDADERO')).toBe(true);
    expect(parseLiteral('Tornillo M8')).toBe('Tornillo M8');
    expect(parseLiteral('  ')).toBe(null);
  });
});

describe('translateFormula', () => {
  it('maps Spanish names and separators but leaves strings alone', () => {
    expect(translateFormula('SUMA(A1:A3; 2)')).toBe('SUM(A1:A3, 2)');
    expect(translateFormula('SI(A1>0;"SI;NO";0)')).toBe('IF(A1>0,"SI;NO",0)');
    expect(translateFormula("'Hoja 2'!A1+redondear(B1;2)")).toBe("'Hoja 2'!A1+ROUND(B1,2)");
  });
});

describe('evaluateWorkbook', () => {
  it('computes chained formulas and ranges', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '2', A2: '3', A3: '=A1*A2', A4: '=SUMA(A1:A3)' }])).values;
    expect(out.s0.A3).toBe(6);
    expect(out.s0.A4).toBe(11);
  });

  it('treats blanks in ranges as empty', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '5', A3: '7', B1: '=SUM(A:A)', B2: '=COUNT(A1:A3)' }])).values;
    expect(out.s0.B1).toBe(12);
    expect(out.s0.B2).toBe(2);
  });

  it('reads other sheets by name', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: "='Hoja 2'!B2*2" }], ['Hoja 2', { B2: '21' }])).values;
    expect(out.s0.A1).toBe(42);
  });

  it('reports errors instead of throwing', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '=1/0', A2: '=A1+1', A3: '=NOEXISTE(1)', A4: '=1+' }])).values;
    expect(out.s0.A1).toEqual({ error: '#DIV/0!' });
    expect(out.s0.A2).toEqual({ error: '#DIV/0!' });
    expect(out.s0.A3).toHaveProperty('error');
    expect(out.s0.A4).toHaveProperty('error');
  });

  it('detects circular references', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '=B1', B1: '=A1' }])).values;
    expect(out.s0.A1).toEqual({ error: '#CIRC!' });
    expect(out.s0.B1).toEqual({ error: '#CIRC!' });
  });
});

describe('Excel compatibility', () => {
  it('keeps apostrophe text and reads scientific notation', () => {
    expect(parseLiteral("'00123")).toBe('00123');
    expect(parseLiteral("'=no es fórmula")).toBe('=no es fórmula');
    expect(parseLiteral('1.5e-7')).toBe(1.5e-7);
  });

  it('resolves named ranges', () => {
    const wb = book(['Hoja 1', { A1: '2', A2: '3', B1: '=SUM(Datos)', B2: '=Primero*10' }]);
    wb.names = { Datos: "'Hoja 1'!$A$1:$A$2", Primero: 'Hoja 1!A1' };
    const out = evaluateWorkbook(wb).values;
    expect(out.s0.B1).toBe(5);
    expect(out.s0.B2).toBe(20);
  });

  it('runs the added functions', () => {
    const out = evaluateWorkbook(
      book([
        'Hoja 1',
        {
          A1: 'Tornillo', A2: 'Tuerca', A3: 'Arandela', A4: '',
          B1: '10', B2: '20', B3: '30',
          C1: '=XLOOKUP("Tuerca",A1:A3,B1:B3)',
          C2: '=XLOOKUP("Clavo",A1:A3,B1:B3,"no hay")',
          C3: '=COUNTA(A1:A4)',
          C4: '=COUNTIFS(A1:A3,"T*",B1:B3,">15")',
          C5: '=UPPER("hola")',
          C6: '=ROUND(PMT(0.01,12,1000),2)',
          C7: '=XMATCH(30,B1:B3)',
        },
      ]),
    ).values;
    expect(out.s0.C1).toBe(20);
    expect(out.s0.C2).toBe('no hay');
    expect(out.s0.C3).toBe(3);
    expect(out.s0.C4).toBe(1);
    expect(out.s0.C5).toBe('HOLA');
    expect(out.s0.C6).toBe(-88.85);
    expect(out.s0.C7).toBe(3);
  });

  it('shows the cached result for formulas it cannot run, and #NAME? otherwise', () => {
    const wb = book(['Hoja 1', { A1: '=UNIQUE(B1:B3)', A2: '=UNIQUE(B1:B3)', A3: '=1/0' }]);
    wb.sheets[0].cells.A1.cached = 'Tornillo';
    wb.sheets[0].cells.A3.cached = 99;
    const { values, approx } = evaluateWorkbook(wb);
    expect(values.s0.A1).toBe('Tornillo');
    expect(approx.has('s0!A1')).toBe(true);
    expect(values.s0.A2).toEqual({ error: '#NAME?' });
    expect(values.s0.A3).toEqual({ error: '#DIV/0!' });
  });
});

describe('formatValue', () => {
  it('uses Spanish number formatting', () => {
    expect(formatValue(1540.5)).toBe('1.540,5');
    expect(formatValue(1540.5, { numFmt: 'currency' })).toBe('$ 1.540,50');
    expect(formatValue(0.15, { numFmt: 'percent' })).toMatch(/^15\s?%$/);
  });
});

describe('format codes', () => {
  it('renders Excel codes with Spanish separators and month names', async () => {
    const { formatWithCode, codeColor } = await import('./format');
    expect(formatWithCode(1234.5, '#,##0.00')).toBe('1.234,50');
    expect(formatWithCode(46313, 'dd/mm/yyyy')).toBe('18/10/2026');
    expect(formatWithCode(46313, 'd mmmm yyyy')).toBe('18 octubre 2026');
    expect(formatWithCode(0.153, '0.0%')).toBe('15,3%');
    expect(codeColor(-5, '#,##0;[Red]-#,##0')).toBe('#ff8f87');
    expect(codeColor(5, '#,##0;[Red]-#,##0')).toBeUndefined();
  });
});
