import { useState } from 'react';
import { rectLabel, rectKeys, type Rect } from '../model/address';
import type { CellFormat, NumFmt, Sheet, Tint, Value, Workbook, WorkbookSummary } from '../model/types';
import { isError } from '../model/types';
import { parseCsv } from '../model/csv';
import type { Edit } from '../state/workbook';
import type { Peer } from '../collab/presence';
import type { CollabState } from '../App';
import type { ImportReport } from '../io/xlsx';
import { supabaseUrl } from '../storage/cloud';
import { Card, PlayIcon, Toggle } from './ui';

export interface ShowOptions {
  colHeads: boolean;
  rowHeads: boolean;
  formulas: boolean;
  compact: boolean;
}

export interface PanelProps {
  wb: Workbook;
  sheet: Sheet;
  values: Record<string, Value>;
  rect: Rect;
  edit(e: Edit): void;
  show: ShowOptions;
  setShow(s: ShowOptions): void;
  chartOn: boolean;
  setChartOn(v: boolean): void;
  recalc(): void;
  insertFunction(name: string): void;
  books: WorkbookSummary[];
  openBook(id: string): void;
  newBook(): void;
  deleteBook(id: string): void;
  exportCsv(): void;
  exportJson(): void;
  openFile(file: File): Promise<void>;
  exportXlsx(): Promise<void>;
  report: ImportReport | null;
  clearReport(): void;
  busy: string | null;
  storeKind: 'local' | 'cloud';
  cloudConfigured: boolean;
  userEmail: string | null;
  signIn(email: string): Promise<string>;
  signOut(): void;
  me: Peer;
  peers: Peer[];
  collab: CollabState;
  isOwner: boolean;
  members: string[];
  invite(email: string): Promise<void>;
  uninvite(email: string): Promise<void>;
}

export function ArchivoPanel(p: PanelProps) {
  const [name, setName] = useState(p.wb.name);
  const [prevName, setPrevName] = useState(p.wb.name);
  if (p.wb.name !== prevName) {
    setPrevName(p.wb.name);
    setName(p.wb.name);
  }
  return (
    <>
      <Card title="Libro">
        <p>
          <label htmlFor="book-name">Nombre</label>
        </p>
        <input
          id="book-name"
          className="field"
          value={name}
          onChange={(e) => setName(e.target.value)}
          onBlur={() => p.edit({ type: 'renameBook', name })}
          onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()}
        />
        <button type="button" className="btn-primary" onClick={p.newBook}>
          + Libro nuevo
        </button>
      </Card>
      <Card title={p.storeKind === 'cloud' ? 'Mis libros (en la nube)' : 'Mis libros (en este navegador)'}>
        <div className="list">
          {p.books.map((b) => (
            <div key={b.id} className="list-item" aria-current={b.id === p.wb.id}>
              <button type="button" className="name" onClick={() => p.openBook(b.id)}>
                {b.name}
              </button>
              <button
                type="button"
                className="icon-btn"
                aria-label={`Borrar ${b.name}`}
                onClick={() => confirm(`¿Borrar "${b.name}"? No se puede deshacer.`) && p.deleteBook(b.id)}
              >
                ×
              </button>
            </div>
          ))}
        </div>
      </Card>
      <Card title="Abrir y guardar">
        <label className="btn-primary" style={{ cursor: 'pointer' }}>
          Abrir archivo de Excel o CSV…
          <input
            type="file"
            accept=".xlsx,.xlsm,.csv,.tsv,text/csv"
            hidden
            onChange={(e) => {
              const file = e.target.files?.[0];
              e.target.value = '';
              if (file) void p.openFile(file);
            }}
          />
        </label>
        <div className="row">
          <button type="button" className="pill" onClick={() => void p.exportXlsx()}>
            Descargar .xlsx
          </button>
          <button type="button" className="pill" onClick={p.exportCsv}>
            Descargar CSV
          </button>
          <button type="button" className="pill" onClick={p.exportJson}>
            Descargar JSON
          </button>
        </div>
        {p.busy && <p>{p.busy}</p>}
        <p>
          Se abre como un libro nuevo. El .xlsx se abre en Excel, LibreOffice, Numbers y Google Sheets con sus fórmulas y
          formatos.
        </p>
      </Card>
      {p.report && <ReportCard report={p.report} onClose={p.clearReport} />}
      <Card title="Desde y hacia Google Sheets">
        <p>
          <b>Traer:</b> en Google Sheets, Archivo → Descargar → Microsoft Excel (.xlsx), y abrilo acá.
          <br />
          <b>Llevar:</b> descargá el .xlsx, subilo a Google Drive y abrilo con Hojas de cálculo de Google.
        </p>
      </Card>
      <Card title="Pegar un CSV en la hoja">
        <label className="pill" style={{ display: 'inline-flex', alignItems: 'center', alignSelf: 'flex-start' }}>
          Importar CSV en la selección…
          <input
            type="file"
            accept=".csv,text/csv,.tsv,text/plain"
            hidden
            onChange={async (e) => {
              const file = e.target.files?.[0];
              e.target.value = '';
              if (!file) return;
              const rows = parseCsv(await file.text());
              p.edit({ type: 'paste', sheetId: p.sheet.id, r: p.rect.r0, c: p.rect.c0, rows });
            }}
          />
        </label>
      </Card>
    </>
  );
}

export function DatosPanel(p: PanelProps) {
  const errors = Object.values(p.values).filter(isError).length;
  const sortable = p.rect.r1 > p.rect.r0;
  return (
    <>
      <Card title="Calcular">
        <button type="button" className="btn-primary" onClick={p.recalc}>
          <PlayIcon />
          Recalcular hoja
        </button>
        {errors === 0 ? (
          <p className="ok">Todas las fórmulas están al día.</p>
        ) : (
          <p className="bad">
            {errors} {errors === 1 ? 'celda tiene' : 'celdas tienen'} un error. Pasá el mouse encima para ver cuál.
          </p>
        )}
      </Card>
      <Card title={`Ordenar ${rectLabel(p.rect)}`}>
        <div className="row">
          <button
            type="button"
            className="pill"
            disabled={!sortable}
            onClick={() => p.edit({ type: 'sortRows', sheetId: p.sheet.id, rect: p.rect, byCol: p.rect.c0, desc: false })}
          >
            A → Z
          </button>
          <button
            type="button"
            className="pill"
            disabled={!sortable}
            onClick={() => p.edit({ type: 'sortRows', sheetId: p.sheet.id, rect: p.rect, byCol: p.rect.c0, desc: true })}
          >
            Z → A
          </button>
        </div>
        <p>{sortable ? 'Ordena las filas seleccionadas por la primera columna.' : 'Seleccioná varias filas para ordenar.'}</p>
      </Card>
      <Card title="Qué mostrar">
        <Toggle on={p.show.colHeads} onChange={(v) => p.setShow({ ...p.show, colHeads: v })}>Letras de columna</Toggle>
        <Toggle on={p.show.rowHeads} onChange={(v) => p.setShow({ ...p.show, rowHeads: v })}>Números de fila</Toggle>
        <Toggle on={p.show.formulas} onChange={(v) => p.setShow({ ...p.show, formulas: v })}>Fórmulas en celdas</Toggle>
        <Toggle on={p.show.compact} onChange={(v) => p.setShow({ ...p.show, compact: v })}>Celdas compactas</Toggle>
      </Card>
      <Card title="Fijar al desplazarse">
        <Toggle
          on={(p.sheet.layout?.frozen?.rows ?? 0) > 0}
          onChange={(v) => p.edit({ type: 'layout', sheetId: p.sheet.id, patch: { frozen: { rows: v ? 1 : 0, cols: p.sheet.layout?.frozen?.cols ?? 0 } } })}
        >
          Primera fila
        </Toggle>
        <Toggle
          on={(p.sheet.layout?.frozen?.cols ?? 0) > 0}
          onChange={(v) => p.edit({ type: 'layout', sheetId: p.sheet.id, patch: { frozen: { rows: p.sheet.layout?.frozen?.rows ?? 0, cols: v ? 1 : 0 } } })}
        >
          Primera columna
        </Toggle>
        <button
          type="button"
          className="pill"
          style={{ alignSelf: 'flex-start' }}
          onClick={() => p.edit({ type: 'layout', sheetId: p.sheet.id, patch: { frozen: { rows: p.rect.r0, cols: p.rect.c0 } } })}
        >
          Fijar hasta la selección
        </button>
      </Card>
    </>
  );
}

const TINTS: { id: Tint; label: string; color: string }[] = [
  { id: 'none', label: 'Sin color', color: '#222d3a' },
  { id: 'blue', label: 'Azul', color: '#1d3047' },
  { id: 'green', label: 'Verde', color: '#1f3a27' },
  { id: 'teal', label: 'Turquesa', color: '#163a40' },
  { id: 'orange', label: 'Naranja', color: '#40261a' },
  { id: 'yellow', label: 'Amarillo', color: '#3d3217' },
  { id: 'red', label: 'Rojo', color: '#4a2326' },
  { id: 'purple', label: 'Violeta', color: '#3a1d4a' },
];

const NUM_FMTS: { id: NumFmt; label: string }[] = [
  { id: 'general', label: 'General' },
  { id: 'number2', label: '0,00' },
  { id: 'currency', label: '$ Moneda' },
  { id: 'percent', label: '% Porcentaje' },
];

export function FormatoPanel(p: PanelProps) {
  const first = p.sheet.cells[rectKeys(p.rect).next().value as string]?.fmt ?? {};
  const apply = (patch: CellFormat) => p.edit({ type: 'format', sheetId: p.sheet.id, rect: p.rect, patch });
  return (
    <>
      <Card title={`Texto · ${rectLabel(p.rect)}`}>
        <div className="row">
          <button type="button" className="pill" aria-pressed={!!first.bold} onClick={() => apply({ bold: !first.bold })} style={{ fontWeight: 700 }}>
            Negrita
          </button>
          <button type="button" className="pill" aria-pressed={!!first.italic} onClick={() => apply({ italic: !first.italic })} style={{ fontStyle: 'italic' }}>
            Cursiva
          </button>
          <button type="button" className="pill" aria-pressed={!!first.underline} onClick={() => apply({ underline: !first.underline })} style={{ textDecoration: 'underline' }}>
            Subrayado
          </button>
          <button type="button" className="pill" aria-pressed={!!first.strike} onClick={() => apply({ strike: !first.strike })} style={{ textDecoration: 'line-through' }}>
            Tachado
          </button>
          <button type="button" className="pill" aria-pressed={!!first.wrap} onClick={() => apply({ wrap: !first.wrap })}>
            Ajustar texto
          </button>
        </div>
        <div className="row">
          {(['left', 'center', 'right'] as const).map((a) => (
            <button key={a} type="button" className="pill" aria-pressed={first.align === a} onClick={() => apply({ align: first.align === a ? undefined : a })}>
              {a === 'left' ? 'Izquierda' : a === 'center' ? 'Centro' : 'Derecha'}
            </button>
          ))}
        </div>
      </Card>
      <Card title="Celdas">
        <div className="row">
          <button type="button" className="pill" disabled={p.rect.r0 === p.rect.r1 && p.rect.c0 === p.rect.c1} onClick={() => p.edit({ type: 'merge', sheetId: p.sheet.id, rect: p.rect })}>
            Combinar
          </button>
          <button type="button" className="pill" onClick={() => p.edit({ type: 'unmerge', sheetId: p.sheet.id, rect: p.rect })}>
            Separar
          </button>
        </div>
        <p>Al combinar queda solo el valor de arriba a la izquierda, como en Excel.</p>
      </Card>
      <Card title="Número">
        <div className="row">
          {NUM_FMTS.map((f) => (
            <button key={f.id} type="button" className="pill" aria-pressed={!first.numFmtCode && (first.numFmt ?? 'general') === f.id} onClick={() => apply({ numFmt: f.id, numFmtCode: undefined })}>
              {f.label}
            </button>
          ))}
        </div>
        {first.numFmtCode && (
          <p>
            Formato del archivo: <span className="mono">{first.numFmtCode}</span>
          </p>
        )}
      </Card>
      <Card title="Color de celda">
        <div className="swatches">
          {TINTS.map((t) => (
            <button
              key={t.id}
              type="button"
              className="swatch"
              aria-label={t.label}
              title={t.label}
              aria-pressed={!first.fill && (first.tint ?? 'none') === t.id}
              style={{ background: t.color }}
              onClick={() => apply({ tint: t.id, fill: undefined })}
            />
          ))}
        </div>
      </Card>
    </>
  );
}

const FUNCTIONS: { name: string; help: string }[] = [
  { name: 'SUMA', help: 'Suma un rango: =SUMA(A1:A10)' },
  { name: 'PROMEDIO', help: 'Promedio: =PROMEDIO(B2:B20)' },
  { name: 'CONTAR', help: 'Cuenta celdas con números' },
  { name: 'MAX', help: 'El valor más grande' },
  { name: 'MIN', help: 'El valor más chico' },
  { name: 'SI', help: '=SI(A1>10; "alto"; "bajo")' },
  { name: 'SUMAR.SI', help: '=SUMAR.SI(B:B; "Acero"; E:E)' },
  { name: 'CONTAR.SI', help: '=CONTAR.SI(G:G; "Activo")' },
  { name: 'BUSCARV', help: '=BUSCARV(valor; A:D; 3; FALSO)' },
  { name: 'REDONDEAR', help: '=REDONDEAR(A1; 2)' },
  { name: 'CONCATENAR', help: 'Une textos' },
  { name: 'HOY', help: 'La fecha de hoy' },
];

export function FormulasPanel(p: PanelProps) {
  return (
    <>
      <Card title="Insertar función">
        <div className="list">
          {FUNCTIONS.map((f) => (
            <div key={f.name} className="list-item">
              <button type="button" className="name" onClick={() => p.insertFunction(f.name)}>
                <span className="mono" style={{ color: 'var(--mode)' }}>{f.name}</span>
                <br />
                <span style={{ fontSize: 12, color: 'var(--muted)' }}>{f.help}</span>
              </button>
            </div>
          ))}
        </div>
      </Card>
      <Card title="Cómo escribir fórmulas">
        <p>
          Empezá con <span className="mono">=</span>. Podés usar nombres en español (SUMA, SI) o en inglés (SUM, IF), y separar
          con <span className="mono">;</span> o <span className="mono">,</span>. Otra hoja: <span className="mono">'Hoja 2'!A1</span>.
          Hay más de 250 funciones disponibles.
        </p>
      </Card>
    </>
  );
}

export function GraficosPanel(p: PanelProps) {
  return (
    <>
      <Card title="Gráfico de barras">
        <Toggle on={p.chartOn} onChange={p.setChartOn}>Mostrar gráfico</Toggle>
        <p>
          Grafica los números de la selección ({rectLabel(p.rect)}). Cada barra toma como nombre el primer texto a su izquierda
          en la misma fila.
        </p>
      </Card>
    </>
  );
}

export function ConectarPanel(p: PanelProps) {
  const [email, setEmail] = useState('');
  const [msg, setMsg] = useState('');
  const [invitee, setInvitee] = useState('');
  const [inviteMsg, setInviteMsg] = useState('');
  const endpoint = `${supabaseUrl}/rest/v1/workbooks?id=eq.${p.wb.id}&select=data`;
  return (
    <>
      <Card title="Edición en vivo">
        {p.collab === 'cloud' && <p className="ok">Conectado. Todo lo que escriban las personas invitadas aparece al instante.</p>}
        {p.collab === 'tabs' && (
          <p>
            Activa entre pestañas de este navegador: abrí este libro en otra pestaña para probarlo. Para editar con otras
            personas, guardá en la nube.
          </p>
        )}
        {p.collab === 'offline' && <p className="bad">No se pudo conectar a la edición en vivo. Tus cambios se siguen guardando.</p>}
        <div className="list">
          <div className="list-item">
            <span className="dot" style={{ background: p.me.color }} />
            <span style={{ fontSize: 13 }}>{p.me.name} (vos)</span>
          </div>
          {p.peers.map((peer) => (
            <div key={peer.id} className="list-item">
              <span className="dot" style={{ background: peer.color }} />
              <span style={{ fontSize: 13 }}>{peer.name}</span>
            </div>
          ))}
        </div>
      </Card>
      {p.isOwner && (
        <Card title="Compartir este libro">
          <form
            className="row"
            onSubmit={async (e) => {
              e.preventDefault();
              try {
                await p.invite(invitee);
                setInviteMsg(`${invitee} ya puede abrirlo entrando con ese email.`);
                setInvitee('');
              } catch {
                setInviteMsg('No se pudo invitar. Probá de nuevo.');
              }
            }}
          >
            <input className="field" type="email" required placeholder="email de la persona" aria-label="Email a invitar" value={invitee} onChange={(e) => setInvitee(e.target.value)} />
            <button type="submit" className="btn-primary" style={{ width: '100%' }}>
              Invitar a editar
            </button>
          </form>
          {inviteMsg && <p>{inviteMsg}</p>}
          <div className="list">
            {p.members.map((m) => (
              <div key={m} className="list-item">
                <span style={{ flex: 1, fontSize: 13 }}>{m}</span>
                <button type="button" className="icon-btn" aria-label={`Quitar a ${m}`} onClick={() => p.uninvite(m)}>
                  ×
                </button>
              </div>
            ))}
          </div>
        </Card>
      )}
      <Card title="Dónde se guarda">
        {p.storeKind === 'cloud' ? (
          <p className="ok">En la nube, con tu cuenta {p.userEmail}. Se sincroniza entre dispositivos.</p>
        ) : (
          <p>En este navegador. Para guardar en la nube y compartir, {p.cloudConfigured ? 'entrá con tu email.' : 'configurá Supabase (ver README).'}</p>
        )}
        {p.cloudConfigured && !p.userEmail && (
          <form
            className="row"
            onSubmit={async (e) => {
              e.preventDefault();
              setMsg(await p.signIn(email));
            }}
          >
            <input className="field" type="email" required placeholder="tu@email.com" aria-label="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
            <button type="submit" className="btn-primary" style={{ width: '100%' }}>
              Enviarme enlace de acceso
            </button>
          </form>
        )}
        {msg && <p>{msg}</p>}
        {p.userEmail && (
          <button type="button" className="pill" onClick={p.signOut}>
            Cerrar sesión
          </button>
        )}
      </Card>
      <Card title="API de este libro">
        {p.storeKind === 'cloud' ? (
          <>
            <div className="field mono" style={{ display: 'flex', alignItems: 'center', fontSize: 12, overflowX: 'auto', whiteSpace: 'nowrap' }}>
              GET {endpoint}
            </div>
            <button type="button" className="pill" onClick={() => navigator.clipboard.writeText(endpoint)}>
              Copiar dirección
            </button>
            <p>
              Otras apps (n8n, Make, Zapier o tu propio código) leen este libro con esa dirección, la clave pública de Supabase y
              un usuario con acceso.
            </p>
          </>
        ) : (
          <p>Disponible cuando el libro se guarda en la nube.</p>
        )}
      </Card>
    </>
  );
}

function ReportCard({ report, onClose }: { report: ImportReport; onClose(): void }) {
  return (
    <section className="card report" aria-label="Resultado de abrir el archivo">
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <h2>Qué se trajo de {report.fileName}</h2>
        <button type="button" className="icon-btn" aria-label="Cerrar" onClick={onClose}>
          ×
        </button>
      </div>
      <p>
        {report.sheets} {report.sheets === 1 ? 'hoja' : 'hojas'}, {report.cells} celdas, {report.formulas} fórmulas.
      </p>
      <h3 className="ok">Funciona igual</h3>
      <ul>
        {report.kept.map((k) => (
          <li key={k}>{k}</li>
        ))}
      </ul>
      {report.preserved.length > 0 && (
        <>
          <h3 style={{ color: 'var(--mode)' }}>Se conserva, con límites</h3>
          <ul>
            {report.preserved.map((k) => (
              <li key={k}>{k}</li>
            ))}
          </ul>
        </>
      )}
      {report.lost.length > 0 && (
        <>
          <h3 className="bad">No se trae todavía</h3>
          <ul>
            {report.lost.map((k) => (
              <li key={k}>{k}</li>
            ))}
          </ul>
          <p>Si guardás desde Cells, estas partes no van a estar en el archivo nuevo. El original no cambia.</p>
        </>
      )}
    </section>
  );
}
