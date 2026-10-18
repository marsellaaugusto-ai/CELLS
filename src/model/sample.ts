import { toKey } from './address';
import type { Cell, Workbook } from './types';
import { newWorkbook } from './types';

/** First-run example so a new user sees formulas, formats and the chart working. */
export function sampleWorkbook(): Workbook {
  const wb = newWorkbook('Inventario de ejemplo');
  const rows = [
    ['Producto', 'Categoría', 'Cantidad', 'Precio', 'Total', 'Estado'],
    ['Tornillo M8', 'Ferretería', '1200', '0,15', '', 'Activo'],
    ['Tuerca M8', 'Ferretería', '1500', '0,08', '', 'Activo'],
    ['Perfil IPN 100', 'Acero', '40', '38,50', '', 'Pedido'],
    ['Cemento 50 kg', 'Obra gris', '85', '9,90', '', 'Activo'],
    ['Arena m³', 'Obra gris', '12', '22', '', 'Bajo stock'],
    ['Varilla 3/8"', 'Acero', '300', '4,20', '', 'Activo'],
    ['Pintura 19 L', 'Acabados', '10', '64', '', 'Activo'],
  ];
  const cells: Record<string, Cell> = {};
  rows.forEach((row, r) =>
    row.forEach((raw, c) => {
      if (r > 0 && c === 4) raw = `=C${r + 1}*D${r + 1}`;
      if (raw) cells[toKey(r, c)] = { raw };
    }),
  );
  for (let c = 0; c < 6; c++) cells[toKey(0, c)] = { ...cells[toKey(0, c)], fmt: { bold: true, tint: 'blue' } };
  for (let r = 1; r < rows.length; r++) {
    cells[toKey(r, 3)].fmt = { numFmt: 'currency' };
    cells[toKey(r, 4)].fmt = { numFmt: 'currency' };
  }
  cells[toKey(5, 5)].fmt = { tint: 'red' };
  const last = rows.length;
  cells[toKey(last + 1, 3)] = { raw: 'Total', fmt: { bold: true, align: 'right' } };
  cells[toKey(last + 1, 4)] = { raw: `=SUMA(E2:E${last})`, fmt: { bold: true, numFmt: 'currency', tint: 'green' } };
  wb.sheets[0].cells = cells;
  return wb;
}
