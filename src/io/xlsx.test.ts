import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { evaluateWorkbook } from '../engine/evaluate';
import { formatValue } from '../engine/format';
import { readXlsx, writeXlsx } from './xlsx';

const fixture = async (name: string) => {
  const buf = await readFile(new URL(`../../test/fixtures/${name}`, import.meta.url));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

describe('readXlsx', () => {
  it('opens a file with charts, styles, merges and validations', async () => {
    const { wb, report } = await readXlsx(await fixture('completo.xlsx'), 'completo.xlsx');
    const [ventas, resumen] = wb.sheets;
    expect(wb.name).toBe('completo');
    expect(ventas.name).toBe('Ventas 2026');
    const c = ventas.cells;

    expect(c.D2.raw).toBe('=B2*C2');
    expect(c.A7.raw).toBe("'00123"); // text that looks like a number stays text
    expect(c.A1.fmt).toMatchObject({ bold: true, color: '#FFFFFF', fill: '#1F4E78', align: 'center', valign: 'middle' });
    expect(c.C2.fmt?.numFmtCode).toBe('"$"#,##0.00');
    expect(c.D5.fmt).toMatchObject({ bold: true, italic: true, underline: true, fontSize: 14 });
    expect(c.D5.fmt?.border?.top).toEqual({ style: 'thin', color: '#FF0000' });
    expect(c.D5.fmt?.border?.bottom?.style).toBe('double');
    expect(c.A8.note).toContain('Revisar antes del cierre');
    expect(c.A9.link).toBe('https://example.com');
    expect(c.F2.validation?.list).toEqual(['Activo', 'Pedido', 'Baja']);
    expect(c.A10.fmt).toMatchObject({ wrap: true, valign: 'top' });

    expect(ventas.layout?.merges).toEqual([{ r0: 11, c0: 0, r1: 12, c1: 2 }]);
    expect(ventas.layout?.frozen).toEqual({ rows: 1, cols: 1 });
    expect(ventas.layout?.colWidths?.[0]).toBeGreaterThan(150);
    expect(ventas.layout?.rowHeights?.[0]).toBe(40);
    expect(ventas.layout?.hiddenCols).toContain(6);
    expect(ventas.layout?.conditional).toHaveLength(1);
    expect(wb.names).toEqual({ Totales: "'Ventas 2026'!$D$2:$D$4" });

    // Formulas calculate, across sheets and through the named range.
    const out = evaluateWorkbook(wb).values;
    expect(out[ventas.id].D5).toBeCloseTo(1200 * 0.15 + 1500 * 0.08 + 40 * 38.5);
    expect(out[resumen.id].B1).toBeCloseTo(1840);
    expect(out[resumen.id].B2).toBe(2400);
    expect(formatValue(out[ventas.id].E2, c.E2.fmt)).toBe('15/01/2026');

    expect(report.lost.join(' ')).toMatch(/1 gráfico/);
    expect(report.preserved.join(' ')).toMatch(/UNIQUE/);
    expect(report.preserved.join(' ')).toMatch(/formato condicional/);
  });

  it('shows the result Excel saved for formulas Cells cannot run', async () => {
    const x = new ExcelJS.Workbook();
    const ws = x.addWorksheet('Datos');
    ws.getCell('A1').value = { formula: '_xlfn.UNIQUE(B1:B3)', result: 'Tornillo' } as ExcelJS.CellFormulaValue;
    const buf = (await x.xlsx.writeBuffer()) as ArrayBuffer;
    const { wb } = await readXlsx(buf, 'dinamico.xlsx');
    const { values, approx } = evaluateWorkbook(wb);
    expect(values[wb.sheets[0].id].A1).toBe('Tornillo');
    expect(approx.size).toBe(1);
  });
});

describe('files from other programs', () => {
  it('opens a file saved by LibreOffice, using its saved results', async () => {
    const { wb, report } = await readXlsx(await fixture('libreoffice.xlsx'), 'libreoffice.xlsx');
    const ventas = wb.sheets[0];
    const { values } = evaluateWorkbook(wb);
    expect(values[ventas.id].D5).toBeCloseTo(1840);
    expect(ventas.cells.A8.note).toContain('Revisar antes del cierre');
    expect(ventas.layout?.merges).toEqual([{ r0: 11, c0: 0, r1: 12, c1: 2 }]);
    expect(report.lost.join(' ')).toMatch(/gráfico/);
  });
});

describe('writeXlsx', () => {
  it('round-trips everything Cells understands', async () => {
    const { wb } = await readXlsx(await fixture('completo.xlsx'), 'completo.xlsx');
    const buf = await writeXlsx(wb, evaluateWorkbook(wb).values);
    const { wb: again } = await readXlsx(buf, 'completo.xlsx');

    const a = wb.sheets[0];
    const b = again.sheets[0];
    expect(b.name).toBe(a.name);
    for (const key of Object.keys(a.cells)) {
      // Validation rules travel as Excel wrote them; ExcelJS may drop default flags, so compare the list.
      const strip = (c: typeof a.cells[string] | undefined) => c && { ...c, cached: undefined, validation: c.validation?.list };
      expect({ key, ...strip(b.cells[key]) }).toEqual({ key, ...strip(a.cells[key]) });
    }
    expect(b.layout?.merges).toEqual(a.layout?.merges);
    expect(b.layout?.frozen).toEqual(a.layout?.frozen);
    expect(b.layout?.colWidths).toEqual(a.layout?.colWidths);
    expect(b.layout?.rowHeights).toEqual(a.layout?.rowHeights);
    expect(b.layout?.hiddenCols).toEqual(a.layout?.hiddenCols);
    expect(b.layout?.conditional).toHaveLength(1);
    expect(again.names).toEqual(wb.names);
    // Excel gets the calculated results too, so it shows numbers before recalculating.
    expect(b.cells.D5.cached).toBeCloseTo(1840);
  });

  it('writes Spanish formulas the way Excel stores them', async () => {
    const { wb } = await readXlsx(await fixture('completo.xlsx'), 'x.xlsx');
    wb.sheets[0].cells.H1 = { raw: '=SUMA(B2:B4; 1)', fmt: { tint: 'green', numFmt: 'currency' } };
    const buf = await writeXlsx(wb, evaluateWorkbook(wb).values);
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(buf);
    const h1 = x.worksheets[0].getCell('H1');
    expect(h1.formula).toBe('SUM(B2:B4, 1)');
    expect(h1.result).toBe(2741);
    expect(h1.numFmt).toBe('"$ "#,##0.00');
    expect((h1.fill as { fgColor?: { argb?: string } }).fgColor?.argb).toBe('FFE2EFDA');
  });
});
