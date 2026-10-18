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
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '2', A2: '3', A3: '=A1*A2', A4: '=SUMA(A1:A3)' }]));
    expect(out.s0.A3).toBe(6);
    expect(out.s0.A4).toBe(11);
  });

  it('treats blanks in ranges as empty', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '5', A3: '7', B1: '=SUM(A:A)', B2: '=COUNT(A1:A3)' }]));
    expect(out.s0.B1).toBe(12);
    expect(out.s0.B2).toBe(2);
  });

  it('reads other sheets by name', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: "='Hoja 2'!B2*2" }], ['Hoja 2', { B2: '21' }]));
    expect(out.s0.A1).toBe(42);
  });

  it('reports errors instead of throwing', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '=1/0', A2: '=A1+1', A3: '=NOEXISTE(1)', A4: '=1+' }]));
    expect(out.s0.A1).toEqual({ error: '#DIV/0!' });
    expect(out.s0.A2).toEqual({ error: '#DIV/0!' });
    expect(out.s0.A3).toHaveProperty('error');
    expect(out.s0.A4).toHaveProperty('error');
  });

  it('detects circular references', () => {
    const out = evaluateWorkbook(book(['Hoja 1', { A1: '=B1', B1: '=A1' }]));
    expect(out.s0.A1).toEqual({ error: '#CIRC!' });
    expect(out.s0.B1).toEqual({ error: '#CIRC!' });
  });
});

describe('formatValue', () => {
  it('uses Spanish number formatting', () => {
    expect(formatValue(1540.5)).toBe('1.540,5');
    expect(formatValue(1540.5, { numFmt: 'currency' })).toBe('$ 1.540,50');
    expect(formatValue(0.15, { numFmt: 'percent' })).toMatch(/^15\s?%$/);
  });
});
