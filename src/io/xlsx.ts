import type ExcelJSNs from 'exceljs';
import type JSZipType from 'jszip';
import { parseKey, toKey } from '../model/address';
import type { BorderSide, BorderStyle, Cell, CellFormat, Range, Sheet, SheetLayout, Tint, Value, Workbook } from '../model/types';
import { isError, newId } from '../model/types';
import { parseLiteral } from '../engine/evaluate';
import { SPANISH_FUNCTIONS, translateFormula } from '../engine/translate';

// Opening and saving Excel files (.xlsx). Google Sheets reads and writes .xlsx too, so this
// is also the bridge to Google Sheets. The libraries are loaded only when a file is used.

export interface ImportReport {
  fileName: string;
  sheets: number;
  cells: number;
  formulas: number;
  /** Things that came over and work. */
  kept: string[];
  /** Things kept in the file (they go back out on export) but not shown or applied yet. */
  preserved: string[];
  /** Things Cells can't carry yet; they are lost if the file is saved from Cells. */
  lost: string[];
}

const load = async () => {
  const [{ default: ExcelJS }, { default: JSZip }] = await Promise.all([import('exceljs'), import('jszip')]);
  return { ExcelJS, JSZip };
};

// ---------- colors ----------

const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

/** Theme colors in the order cells refer to them: lt1, dk1, lt2, dk2, accent1…6. */
function themePalette(themeXml: string | undefined): string[] {
  if (!themeXml) return DEFAULT_THEME;
  const pick = (tag: string) => {
    const m = new RegExp(`<a:${tag}>\\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")`).exec(themeXml);
    return m ? (m[1] ?? m[2]).toUpperCase() : null;
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  return order.map((t, i) => pick(t) ?? DEFAULT_THEME[i]);
}

function applyTint(hex: string, tint: number) {
  const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = ch.map((c) => Math.round(tint < 0 ? c * (1 + tint) : c + (255 - c) * tint));
  return out.map((c) => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}

type XColor = { argb?: string; theme?: number; tint?: number; indexed?: number } | undefined;

const INDEXED: Record<number, string> = { 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00', 12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF' };

function colorOf(c: XColor, theme: string[]): string | undefined {
  if (!c) return undefined;
  let hex: string | undefined;
  if (c.argb) hex = c.argb.slice(-6).toUpperCase();
  else if (c.theme !== undefined) hex = theme[c.theme];
  else if (c.indexed !== undefined) hex = INDEXED[c.indexed];
  if (!hex) return undefined;
  if (c.tint) hex = applyTint(hex, c.tint);
  return `#${hex}`;
}

// ---------- zip fixes and feature detection ----------

/**
 * Some programs write internal paths as absolute ("/xl/drawings/drawing1.xml"). ExcelJS only
 * understands relative ones and crashes otherwise, so rewrite them before loading.
 */
async function normalizePackage(zip: JSZipType) {
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith('.rels')) continue;
    const baseDir = path.replace(/_rels\/[^/]+$/, '').replace(/\/$/, '');
    const xml = await zip.file(path)!.async('string');
    const fixed = xml.replace(/Target="\/([^"]+)"/g, (_m, target: string) => {
      const from = baseDir ? baseDir.split('/') : [];
      const to = target.split('/');
      let i = 0;
      while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
      return `Target="${[...Array(from.length - i).fill('..'), ...to.slice(i)].join('/')}"`;
    });
    if (fixed !== xml) zip.file(path, fixed);
  }
  // ExcelJS also only finds notes at the file names Excel uses (xl/comments1.xml and
  // xl/drawings/vmlDrawing1.vml); other programs name them differently, so rename them.
  let n = 0;
  for (const relsPath of Object.keys(zip.files).filter((p) => /^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(p))) {
    let xml = await zip.file(relsPath)!.async('string');
    const renames: [string, RegExp, (k: number) => string][] = [
      ['/comments"', /^xl\/comments\d+\.xml$/, (k) => `xl/comments${k}.xml`],
      ['/vmlDrawing"', /^xl\/drawings\/vmlDrawing\d+\.vml$/, (k) => `xl/drawings/vmlDrawing${k}.vml`],
    ];
    for (const [typeEnd, ok, rename] of renames) {
      xml = xml.replace(/<Relationship\b[^>]*>/g, (rel) => {
        if (!rel.includes(typeEnd)) return rel;
        const target = /Target="([^"]+)"/.exec(rel)?.[1];
        if (!target) return rel;
        const resolved = new URL(target, 'file:///xl/worksheets/').pathname.slice(1);
        if (ok.test(resolved) || !zip.file(resolved)) return rel;
        const next = rename(1000 + ++n);
        zip.file(next, zip.file(resolved)!.async('uint8array'));
        zip.remove(resolved);
        return rel.replace(/Target="[^"]+"/, `Target="../${next.slice(3)}"`);
      });
    }
    zip.file(relsPath, xml);
  }
  // Notes written as plain <t> (instead of Excel's <r><t> runs) come out empty in ExcelJS.
  for (const path of Object.keys(zip.files).filter((p) => /^xl\/comments\d+\.xml$/.test(p))) {
    const xml = await zip.file(path)!.async('string');
    const fixed = xml.replace(/<text>\s*<t([^>]*)>([\s\S]*?)<\/t>\s*<\/text>/g, '<text><r><t$1>$2</t></r></text>');
    if (fixed !== xml) zip.file(path, fixed);
  }
  // ExcelJS only reads drawings written the way Excel writes them ("xdr:" prefix) and crashes
  // on charts it doesn't know. Charts aren't imported anyway (they're counted for the report
  // first), so any other drawing is replaced by an empty one; its pictures are counted as lost.
  let droppedPictures = 0;
  for (const path of Object.keys(zip.files)) {
    if (!/^xl\/drawings\/[^/]+\.xml$/.test(path)) continue;
    const xml = await zip.file(path)!.async('string');
    const pictures = (xml.match(/<(?:\w+:)?pic>/g) ?? []).length;
    if (pictures && xml.includes('<xdr:wsDr')) continue;
    droppedPictures += pictures;
    zip.file(
      path,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>',
    );
  }
  return { droppedPictures };
}

function countEntries(zip: JSZipType, re: RegExp) {
  return Object.keys(zip.files).filter((p) => re.test(p)).length;
}

const plural = (n: number, one: string, many: string) => `${n} ${n === 1 ? one : many}`;

// ---------- import ----------

const BORDER_STYLES: Record<string, BorderStyle> = {
  thin: 'thin', medium: 'medium', thick: 'thick', dashed: 'dashed', dotted: 'dotted', double: 'double', hair: 'hair',
  mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed', dashDotDot: 'dotted', mediumDashDotDot: 'dotted', slantDashDot: 'dashed',
};

const DEFAULT_FONTS = new Set(['calibri', 'aptos', 'aptos narrow', 'arial']);

function excelSerial(d: Date): number {
  return (d.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000;
}

/** Engine-supported function names, to report formulas that will rely on Excel's last result. */
let knownFunctions: Set<string> | null = null;
async function supportedFunctions() {
  if (knownFunctions) return knownFunctions;
  const { default: FormulaParser } = await import('fast-formula-parser');
  const { EXTRA_FUNCTIONS } = await import('../engine/functions');
  const p = new FormulaParser({});
  knownFunctions = new Set([
    ...Object.keys(p.functions ?? {}),
    ...Object.keys(p.funsNeedContext ?? {}),
    ...Object.keys(p.funsPreserveRef ?? {}),
    ...Object.keys(EXTRA_FUNCTIONS),
  ]);
  return knownFunctions;
}

/** Text that would otherwise be read as a number, date, boolean or formula gets the ' prefix. */
function textRaw(s: string) {
  return s.startsWith('=') || s.startsWith("'") || typeof parseLiteral(s) !== 'string' ? `'${s}` : s;
}

function cachedOf(result: unknown): Cell['cached'] {
  if (result instanceof Date) return excelSerial(result);
  if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') return result;
  return undefined;
}

export async function readXlsx(data: ArrayBuffer, fileName: string): Promise<{ wb: Workbook; report: ImportReport }> {
  const { ExcelJS, JSZip } = await load();
  const zip = await JSZip.loadAsync(data);
  const { droppedPictures } = await normalizePackage(zip);
  const themeXml = await zip.file('xl/theme/theme1.xml')?.async('string');
  const theme = themePalette(themeXml);

  const x = new ExcelJS.Workbook();
  await x.xlsx.load(await zip.generateAsync({ type: 'arraybuffer' }));

  const known = await supportedFunctions();
  const unknownFns = new Map<string, number>();
  const counts = { cells: 0, formulas: 0, rich: 0, notes: 0, links: 0, validations: 0, merges: 0, conditional: 0, names: 0, images: 0, tables: 0 };
  let otherValidations = 0;
  let frozenSheets = 0;
  let hiddenSheets = 0;

  const sheets: Sheet[] = [];
  x.eachSheet((ws) => {
    if (ws.state && ws.state !== 'visible') hiddenSheets++;
    const cells: Record<string, Cell> = {};
    const layout: SheetLayout = {};

    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (row.height && (row as unknown as { customHeight?: boolean }).customHeight !== false)
        (layout.rowHeights ??= {})[rowNumber - 1] = Math.round((row.height * 4) / 3);
      if (row.hidden) (layout.hiddenRows ??= []).push(rowNumber - 1);
      row.eachCell({ includeEmpty: false }, (xc) => {
        if (xc.type === ExcelJS.ValueType.Merge) return;
        const key = toKey(Number(xc.row) - 1, Number(xc.col) - 1);
        const cell: Cell = { raw: '' };
        const v = xc.value as unknown;
        switch (xc.type) {
          case ExcelJS.ValueType.Formula: {
            const f = (xc.formula ?? (v as { formula?: string }).formula ?? '') as string;
            cell.raw = '=' + f;
            const cached = cachedOf((v as { result?: unknown }).result);
            if (cached !== undefined) cell.cached = cached;
            counts.formulas++;
            for (const m of f.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) {
              const name = m[1].toUpperCase().replace(/^_XLFN\.|^_XLWS\./, '');
              if (!known.has(name) && !SPANISH_FUNCTIONS[name]) unknownFns.set(name, (unknownFns.get(name) ?? 0) + 1);
            }
            break;
          }
          case ExcelJS.ValueType.Number:
            cell.raw = String(v);
            break;
          case ExcelJS.ValueType.Boolean:
            cell.raw = v ? 'TRUE' : 'FALSE';
            break;
          case ExcelJS.ValueType.Date:
            cell.raw = String(excelSerial(v as Date));
            break;
          case ExcelJS.ValueType.Error:
            cell.raw = textRaw((v as { error: string }).error);
            break;
          case ExcelJS.ValueType.Hyperlink: {
            const h = v as { text?: unknown; hyperlink: string };
            const text = typeof h.text === 'object' && h.text && 'richText' in h.text
              ? (h.text as { richText: { text: string }[] }).richText.map((t) => t.text).join('')
              : String(h.text ?? h.hyperlink);
            cell.raw = textRaw(text);
            cell.link = h.hyperlink;
            counts.links++;
            break;
          }
          case ExcelJS.ValueType.RichText:
            cell.raw = textRaw((v as { richText: { text: string }[] }).richText.map((t) => t.text).join(''));
            counts.rich++;
            break;
          default:
            cell.raw = v === null || v === undefined ? '' : textRaw(String(v));
        }

        const fmt = formatFrom(xc, theme);
        if (fmt) cell.fmt = fmt;
        // Dates need a date format to read as dates; Excel always stores one, but be safe.
        if (xc.type === ExcelJS.ValueType.Date && !cell.fmt?.numFmtCode) cell.fmt = { ...cell.fmt, numFmtCode: 'dd/mm/yyyy' };

        const note = xc.note;
        const noteText = typeof note === 'string' ? note : (note?.texts ?? []).map((t) => t.text).join('');
        if (noteText) {
          cell.note = noteText;
          counts.notes++;
        }
        const dv = xc.dataValidation as { type?: string; formulae?: unknown[] } | undefined;
        if (dv?.type) {
          const f = String(dv.formulae?.[0] ?? '');
          const list = dv.type === 'list' && /^".*"$/.test(f) ? f.slice(1, -1).split(',').map((s) => s.trim()) : null;
          if (!list) otherValidations++;
          cell.validation = { list, excel: JSON.parse(JSON.stringify(dv)) };
          counts.validations++;
        }
        if (cell.raw === '' && !cell.fmt && !cell.note && !cell.validation) return;
        cells[key] = cell;
        counts.cells++;
      });
    });

    const columns = (ws.columns ?? []) as ExcelJSNs.Column[];
    columns.forEach((col, i) => {
      if (col.width && Math.abs(col.width - 9.140625) > 0.01) (layout.colWidths ??= {})[i] = Math.round(col.width * 7 + 5);
      if (col.hidden) (layout.hiddenCols ??= []).push(i);
    });
    const merges = ((ws.model as { merges?: string[] }).merges ?? []).map((m): Range => {
      const [a, b] = m.split(':');
      const p = parseKey(a);
      const q = parseKey(b ?? a);
      return { r0: p.r, c0: p.c, r1: q.r, c1: q.c };
    });
    if (merges.length) {
      layout.merges = merges;
      counts.merges += merges.length;
    }
    const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined;
    if (view?.state === 'frozen' && (view.xSplit || view.ySplit)) {
      layout.frozen = { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 };
      frozenSheets++;
    }
    const cf = (ws as unknown as { conditionalFormattings?: unknown[] }).conditionalFormattings ?? [];
    if (cf.length) {
      layout.conditional = JSON.parse(JSON.stringify(cf));
      counts.conditional += cf.length;
    }
    counts.images += ws.getImages().length;
    counts.tables += Object.keys((ws as unknown as { tables?: object }).tables ?? {}).length;

    const used = Object.keys(cells).map(parseKey);
    sheets.push({
      id: newId(),
      name: ws.name,
      cells,
      rows: Math.max(100, ...used.map((p) => p.r + 21)),
      cols: Math.max(26, ...used.map((p) => p.c + 6)),
      ...(Object.keys(layout).length ? { layout } : {}),
    });
  });

  const names: Record<string, string> = {};
  for (const n of (x.definedNames.model ?? []) as { name: string; ranges: string[] }[]) {
    if (n.ranges?.[0]) names[n.name] = n.ranges[0];
  }
  counts.names = Object.keys(names).length;

  const wb: Workbook = {
    id: newId(),
    name: fileName.replace(/\.(xlsx|xlsm)$/i, ''),
    sheets: sheets.length ? sheets : [{ id: newId(), name: 'Hoja 1', cells: {}, rows: 100, cols: 26 }],
    updatedAt: new Date().toISOString(),
    ...(counts.names ? { names } : {}),
  };

  const kept = [
    'Valores, textos, fechas y fórmulas',
    'Formato de números, fuentes, colores, bordes y alineación',
    'Ancho de columnas, alto de filas, filas y columnas ocultas',
  ];
  if (counts.merges) kept.push(plural(counts.merges, 'grupo de celdas combinadas', 'grupos de celdas combinadas'));
  if (frozenSheets) kept.push('Filas y columnas fijas');
  if (counts.notes) kept.push(plural(counts.notes, 'nota', 'notas'));
  if (counts.links) kept.push(plural(counts.links, 'enlace', 'enlaces'));
  if (counts.names) kept.push(plural(counts.names, 'rango con nombre', 'rangos con nombre'));
  if (counts.validations - otherValidations > 0) kept.push('Listas desplegables (validación de datos)');

  const preserved: string[] = [];
  if (counts.conditional) preserved.push(`${plural(counts.conditional, 'regla', 'reglas')} de formato condicional (se guardan; todavía no se ven)`);
  if (otherValidations) preserved.push(`${plural(otherValidations, 'regla', 'reglas')} de validación que no son listas (se guardan; todavía no se aplican)`);
  if (unknownFns.size) {
    const list = [...unknownFns.entries()].map(([n, c]) => `${n} (${c})`).join(', ');
    preserved.push(`Fórmulas con funciones que Cells todavía no calcula: ${list}. Muestran el último resultado de Excel.`);
  }
  if (hiddenSheets) preserved.push(`${plural(hiddenSheets, 'hoja oculta', 'hojas ocultas')} (se muestran)`);

  const lost: string[] = [];
  const charts = countEntries(zip, /^xl\/charts\/chart\d+\.xml$/);
  const pivots = countEntries(zip, /^xl\/pivotTables\/.+\.xml$/);
  if (charts) lost.push(plural(charts, 'gráfico', 'gráficos'));
  if (pivots) lost.push(plural(pivots, 'tabla dinámica', 'tablas dinámicas') + ' (quedan sus valores)');
  if (counts.images + droppedPictures) lost.push(plural(counts.images + droppedPictures, 'imagen', 'imágenes'));
  if (counts.tables) lost.push(plural(counts.tables, 'tabla con formato', 'tablas con formato') + ' (quedan sus celdas)');
  if (zip.file('xl/vbaProject.bin')) lost.push('Macros (VBA)');
  if (countEntries(zip, /^xl\/slicers\//)) lost.push('Segmentaciones de datos');
  if (countEntries(zip, /^xl\/externalLinks\//)) lost.push('Vínculos a otros archivos (quedan los últimos valores)');
  if (counts.rich) lost.push(`${plural(counts.rich, 'celda', 'celdas')} con varios formatos dentro del texto (queda el texto)`);

  return {
    wb,
    report: { fileName, sheets: wb.sheets.length, cells: counts.cells, formulas: counts.formulas, kept, preserved, lost },
  };
}

function formatFrom(xc: ExcelJSNs.Cell, theme: string[]): CellFormat | undefined {
  const f: CellFormat = {};
  const font = xc.font ?? {};
  if (font.bold) f.bold = true;
  if (font.italic) f.italic = true;
  if (font.underline) f.underline = true;
  if (font.strike) f.strike = true;
  if (font.size && font.size !== 11 && font.size !== 10) f.fontSize = font.size;
  if (font.name && !DEFAULT_FONTS.has(font.name.toLowerCase())) f.fontFamily = font.name;
  const color = colorOf(font.color as XColor, theme);
  if (color && color !== '#000000' && !((font.color as XColor)?.theme === 1 && !(font.color as XColor)?.tint)) f.color = color;

  const fill = xc.fill as { type?: string; pattern?: string; fgColor?: XColor } | undefined;
  if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const bg = colorOf(fill.fgColor, theme);
    if (bg) f.fill = bg;
  }

  const al = xc.alignment ?? {};
  if (al.horizontal === 'left' || al.horizontal === 'center' || al.horizontal === 'right') f.align = al.horizontal;
  else if (al.horizontal === 'centerContinuous') f.align = 'center';
  if (al.vertical === 'top' || al.vertical === 'bottom') f.valign = al.vertical;
  else if (al.vertical === 'middle') f.valign = 'middle';
  if (al.wrapText) f.wrap = true;

  const border: NonNullable<CellFormat['border']> = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const b = (xc.border as Record<string, { style?: string; color?: XColor }> | undefined)?.[side];
    if (b?.style && BORDER_STYLES[b.style]) {
      const s: BorderSide = { style: BORDER_STYLES[b.style] };
      const c = colorOf(b.color, theme);
      if (c) s.color = c;
      border[side] = s;
    }
  }
  if (Object.keys(border).length) f.border = border;

  if (xc.numFmt && xc.numFmt !== 'General') f.numFmtCode = xc.numFmt;
  return Object.keys(f).length ? f : undefined;
}

// ---------- export ----------

/** Light versions of the preset tints, so exported files read well on Excel's white sheets. */
const TINT_EXPORT: Record<Exclude<Tint, 'none'>, string> = {
  blue: 'DDEBF7', green: 'E2EFDA', teal: 'DDF3F5', orange: 'FCE4D6', yellow: 'FFF2CC', red: 'F8D7DA', purple: 'EADCF2',
};
const PRESET_CODES = { number2: '#,##0.00', currency: '"$ "#,##0.00', percent: '0.00%' } as const;

const argb = (hex: string) => 'FF' + hex.replace('#', '').toUpperCase();

function toExcelValue(v: Value): unknown {
  if (v === null) return null;
  if (isError(v)) return { error: v.error };
  return v;
}

export async function writeXlsx(wb: Workbook, values: Record<string, Record<string, Value>>): Promise<ArrayBuffer> {
  const { ExcelJS } = await load();
  const x = new ExcelJS.Workbook();
  x.creator = 'Cells';

  for (const sheet of wb.sheets) {
    const ws = x.addWorksheet(sheet.name.slice(0, 31));
    const layout = sheet.layout ?? {};
    for (const [c, px] of Object.entries(layout.colWidths ?? {})) ws.getColumn(Number(c) + 1).width = Math.max(0, (px - 5) / 7);
    for (const c of layout.hiddenCols ?? []) ws.getColumn(c + 1).hidden = true;
    for (const [r, px] of Object.entries(layout.rowHeights ?? {})) ws.getRow(Number(r) + 1).height = (px * 3) / 4;
    for (const r of layout.hiddenRows ?? []) ws.getRow(r + 1).hidden = true;
    if (layout.frozen) ws.views = [{ state: 'frozen', xSplit: layout.frozen.cols, ySplit: layout.frozen.rows }];

    const computed = values[sheet.id] ?? {};
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const xc = ws.getCell(key);
      const raw = cell.raw;
      if (raw.length > 1 && raw.startsWith('=')) {
        const result = computed[key] ?? null;
        xc.value = { formula: translateFormula(raw.slice(1)), result: toExcelValue(result) } as ExcelJSNs.CellFormulaValue;
      } else if (raw !== '') {
        const v = parseLiteral(raw);
        xc.value = cell.link ? { text: String(v), hyperlink: cell.link } : (v as ExcelJSNs.CellValue);
      }
      if (cell.note) xc.note = cell.note;
      if (cell.validation) {
        xc.dataValidation = (cell.validation.excel as ExcelJSNs.DataValidation) ?? {
          type: 'list', allowBlank: true, formulae: [`"${(cell.validation.list ?? []).join(',')}"`],
        };
      }
      const f = cell.fmt;
      if (!f) continue;
      const font: Partial<ExcelJSNs.Font> = {};
      if (f.bold) font.bold = true;
      if (f.italic) font.italic = true;
      if (f.underline) font.underline = true;
      if (f.strike) font.strike = true;
      if (f.fontSize) font.size = f.fontSize;
      if (f.fontFamily) font.name = f.fontFamily;
      if (f.color) font.color = { argb: argb(f.color) };
      if (Object.keys(font).length) xc.font = font;
      const bg = f.fill ? argb(f.fill) : f.tint && f.tint !== 'none' ? 'FF' + TINT_EXPORT[f.tint] : null;
      if (bg) xc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      const alignment: Partial<ExcelJSNs.Alignment> = {};
      if (f.align) alignment.horizontal = f.align;
      if (f.valign) alignment.vertical = f.valign;
      if (f.wrap) alignment.wrapText = true;
      if (Object.keys(alignment).length) xc.alignment = alignment;
      if (f.border) {
        const border: Partial<ExcelJSNs.Borders> = {};
        for (const [side, b] of Object.entries(f.border)) {
          if (b) (border as Record<string, unknown>)[side] = { style: b.style, ...(b.color ? { color: { argb: argb(b.color) } } : {}) };
        }
        xc.border = border;
      }
      const code = f.numFmtCode ?? (f.numFmt && f.numFmt !== 'general' ? PRESET_CODES[f.numFmt] : undefined);
      if (code) xc.numFmt = code;
    }
    for (const m of layout.merges ?? []) ws.mergeCells(m.r0 + 1, m.c0 + 1, m.r1 + 1, m.c1 + 1);
    for (const rule of layout.conditional ?? []) {
      try {
        ws.addConditionalFormatting(rule as ExcelJSNs.ConditionalFormattingOptions);
      } catch {
        // A rule ExcelJS can't write back is skipped rather than failing the whole export.
      }
    }
  }
  for (const [name, ref] of Object.entries(wb.names ?? {})) {
    try {
      x.definedNames.add(ref, name);
    } catch {
      // ignore names ExcelJS rejects
    }
  }
  return (await x.xlsx.writeBuffer()) as ArrayBuffer;
}
