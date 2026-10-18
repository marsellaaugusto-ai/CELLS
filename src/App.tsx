import { useCallback, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { rectOf, rectLabel, toKey } from './model/address';
import { download, parseDelimited, toCsv } from './model/csv';
import { sampleWorkbook } from './model/sample';
import type { Workbook, WorkbookSummary } from './model/types';
import { isError, newWorkbook } from './model/types';
import { evaluateWorkbook } from './engine/evaluate';
import { rectToTsv, usedRect, useWorkbook } from './state/workbook';
import { localStore } from './storage/local';
import { cloudStore, supabase } from './storage/cloud';
import type { Store } from './storage/store';
import { Chart } from './components/Chart';
import { Grid, type GridActions, type Selection } from './components/Grid';
import {
  ArchivoPanel,
  ConectarPanel,
  DatosPanel,
  FormatoPanel,
  FormulasPanel,
  GraficosPanel,
  type PanelProps,
  type ShowOptions,
} from './components/Panels';
import { MODES, ModeHead, Rail, type ModeId } from './components/ui';

type SaveState = 'saved' | 'saving' | 'error';

const PANELS = {
  archivo: ArchivoPanel,
  datos: DatosPanel,
  formato: FormatoPanel,
  formulas: FormulasPanel,
  graficos: GraficosPanel,
  conectar: ConectarPanel,
};

export default function App() {
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [authChecked, setAuthChecked] = useState(!supabase);
  const store: Store = useMemo(() => (supabase && userEmail ? cloudStore(supabase) : localStore), [userEmail]);

  useEffect(() => {
    if (!supabase) return;
    supabase.auth.getSession().then(({ data }) => {
      setUserEmail(data.session?.user.email ?? null);
      setAuthChecked(true);
    });
    const { data } = supabase.auth.onAuthStateChange((_e, session) => setUserEmail(session?.user.email ?? null));
    return () => data.subscription.unsubscribe();
  }, []);

  const { wb, edit, load, undo, redo } = useWorkbook(newWorkbook());
  const [sheetId, setSheetId] = useState(wb.sheets[0].id);
  const sheet = wb.sheets.find((s) => s.id === sheetId) ?? wb.sheets[0];
  const [sel, setSel] = useState<Selection>({ anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } });
  const [editing, setEditing] = useState<{ initial: string } | null>(null);
  const [mode, setMode] = useState<ModeId>('datos');
  const [show, setShow] = useState<ShowOptions>({ colHeads: true, rowHeads: true, formulas: false, compact: false });
  const [chartOn, setChartOn] = useState(false);
  const [recalcTick, setRecalcTick] = useState(0);
  const [renaming, setRenaming] = useState<string | null>(null);

  const [ready, setReady] = useState(false);
  const [books, setBooks] = useState<WorkbookSummary[]>([]);
  const [saveState, setSaveState] = useState<SaveState>('saved');
  const loadedRef = useRef<Workbook | null>(null);
  const wbRef = useRef(wb);
  wbRef.current = wb;

  const refreshBooks = useCallback(() => store.list().then(setBooks).catch(() => setBooks([])), [store]);

  const open = useCallback(
    (next: Workbook) => {
      loadedRef.current = next;
      load(next);
      setSheetId(next.sheets[0].id);
      setSel({ anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } });
      setEditing(null);
    },
    [load],
  );

  // Boot (and re-boot when signing in/out): open the most recent workbook, or create the example.
  useEffect(() => {
    if (!authChecked) return;
    let cancelled = false;
    (async () => {
      const list = await store.list().catch(() => []);
      if (cancelled) return;
      let first = list[0] ? await store.load(list[0].id).catch(() => null) : null;
      if (!first) {
        first = sampleWorkbook();
        await store.save(first).catch(() => setSaveState('error'));
      }
      if (cancelled) return;
      open(first);
      setReady(true);
      refreshBooks();
    })();
    return () => {
      cancelled = true;
    };
  }, [store, authChecked, open, refreshBooks]);

  // Autosave shortly after each change.
  useEffect(() => {
    if (!ready || wb === loadedRef.current) return;
    setSaveState('saving');
    const t = setTimeout(() => {
      store
        .save(wb)
        .then(() => {
          setSaveState('saved');
          refreshBooks();
        })
        .catch(() => setSaveState('error'));
    }, 600);
    return () => clearTimeout(t);
  }, [wb, ready, store, refreshBooks]);

  // Live updates when another tab, device or app saves this workbook.
  useEffect(() => {
    if (!ready) return;
    return store.subscribe(wb.id, (remote) => {
      if (remote.updatedAt > wbRef.current.updatedAt) {
        loadedRef.current = remote;
        load(remote);
      }
    });
  }, [store, wb.id, ready, load]);

  const computed = useMemo(() => evaluateWorkbook(wb), [wb, recalcTick]);
  const values = computed[sheet.id] ?? {};
  const rect = rectOf(sel.anchor, sel.focus);
  const activeKey = toKey(sel.focus.r, sel.focus.c);
  const activeRaw = sheet.cells[activeKey]?.raw ?? '';

  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);

  const actions: GridActions = {
    select: (s) => {
      setSel(s);
      setFormulaDraft(null);
    },
    startEdit: (initial) => setEditing({ initial: initial ?? activeRaw }),
    commitEdit: (raw, move) => {
      if (raw !== activeRaw) edit({ type: 'setCell', sheetId: sheet.id, key: activeKey, raw });
      setEditing(null);
      const f = sel.focus;
      const next =
        move === 'down' ? { r: Math.min(f.r + 1, sheet.rows - 1), c: f.c }
        : move === 'right' ? { r: f.r, c: Math.min(f.c + 1, sheet.cols - 1) }
        : f;
      setSel({ anchor: next, focus: next });
    },
    cancelEdit: () => setEditing(null),
    clear: () => edit({ type: 'clear', sheetId: sheet.id, rect }),
    undo,
    redo,
    toggleBold: () => {
      const bold = !sheet.cells[activeKey]?.fmt?.bold;
      edit({ type: 'format', sheetId: sheet.id, rect, patch: { bold } });
    },
    addRows: () => edit({ type: 'addRows', sheetId: sheet.id, count: 50 }),
  };

  const onCopy = (cut: boolean) => {
    const text = rectToTsv(sheet, rect);
    if (cut) edit({ type: 'clear', sheetId: sheet.id, rect });
    return text;
  };

  const onPaste = (text: string) => {
    const rows = parseDelimited(text.replace(/\r?\n$/, ''), '\t');
    if (rows.length) edit({ type: 'paste', sheetId: sheet.id, r: rect.r0, c: rect.c0, rows });
  };

  const exportCsv = () => {
    const used = usedRect(sheet);
    if (!used) return;
    const rows: string[][] = [];
    for (let r = 0; r <= used.r1; r++) {
      const row: string[] = [];
      for (let c = 0; c <= used.c1; c++) {
        const v = values[toKey(r, c)] ?? null;
        row.push(v === null ? '' : isError(v) ? v.error : String(v));
      }
      rows.push(row);
    }
    download(`${wb.name} - ${sheet.name}.csv`, toCsv(rows), 'text/csv;charset=utf-8');
  };

  const signIn = async (email: string) => {
    if (!supabase) return 'La nube no está configurada.';
    const { error } = await supabase.auth.signInWithOtp({ email, options: { emailRedirectTo: location.href } });
    return error ? `No se pudo enviar: ${error.message}` : 'Listo: revisá tu email y abrí el enlace.';
  };

  const panelProps: PanelProps = {
    wb,
    sheet,
    values,
    rect,
    edit,
    show,
    setShow,
    chartOn,
    setChartOn,
    recalc: () => setRecalcTick((t) => t + 1),
    insertFunction: (name) => setEditing({ initial: `=${name}(` }),
    books,
    openBook: async (id) => {
      const next = await store.load(id).catch(() => null);
      if (next) open(next);
    },
    newBook: async () => {
      const next = newWorkbook();
      await store.save(next).catch(() => setSaveState('error'));
      open(next);
      refreshBooks();
    },
    deleteBook: async (id) => {
      await store.remove(id).catch(() => setSaveState('error'));
      const list = await store.list().catch(() => []);
      setBooks(list);
      if (id === wb.id) {
        const next = (list[0] && (await store.load(list[0].id).catch(() => null))) || newWorkbook();
        if (!list[0]) await store.save(next).catch(() => setSaveState('error'));
        open(next);
        refreshBooks();
      }
    },
    exportCsv,
    exportJson: () => download(`${wb.name}.cells.json`, JSON.stringify(wb, null, 2), 'application/json'),
    storeKind: store.kind,
    cloudConfigured: !!supabase,
    userEmail,
    signIn,
    signOut: () => void supabase?.auth.signOut(),
  };

  const modeIndex = MODES.findIndex((m) => m.id === mode);
  const m = MODES[modeIndex];
  const Panel = PANELS[mode];
  const appStyle = {
    '--mode': m.color,
    '--mode-strong': m.strong,
    '--mode-tint': m.tint,
    '--gap': show.compact ? '3px' : '6px',
    '--row-h': show.compact ? '26px' : '32px',
  } as CSSProperties;

  const formulaValue = formulaDraft ?? activeRaw;

  return (
    <div className="app" style={appStyle}>
      <Rail
        mode={mode}
        onMode={(id) => {
          setMode(id);
          if (id === 'graficos') setChartOn(true);
        }}
      />

      <aside className="panel" aria-label={`Panel ${m.label}`}>
        <ModeHead index={modeIndex + 1} label={m.label} />
        <Panel {...panelProps} />
      </aside>

      <main className="main">
        <section className="display" aria-label="Barra de fórmula">
          <div className="display-strip">
            <span>Fórmula</span>
            <span>{rectLabel(rect)}</span>
          </div>
          <div className="formula">
            <span className="addr">{activeKey}</span>
            <input
              aria-label={`Contenido de ${activeKey}`}
              value={formulaValue}
              placeholder="Escribí un valor o una fórmula que empiece con ="
              onChange={(e) => setFormulaDraft(e.target.value)}
              onBlur={() => {
                if (formulaDraft !== null && formulaDraft !== activeRaw)
                  edit({ type: 'setCell', sheetId: sheet.id, key: activeKey, raw: formulaDraft });
                setFormulaDraft(null);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
                if (e.key === 'Escape') {
                  setFormulaDraft(null);
                  (e.target as HTMLInputElement).blur();
                }
              }}
            />
          </div>
        </section>

        <section className="display" style={{ flex: 1 }} aria-label="Hoja">
          <div className="display-strip">
            <span>
              {wb.name} · {sheet.name}
            </span>
            <span>{ready ? '' : 'Cargando…'}</span>
          </div>
          <Grid
            sheet={sheet}
            values={values}
            sel={sel}
            editing={editing}
            show={show}
            actions={actions}
            onCopy={onCopy}
            onPaste={onPaste}
          />
        </section>

        {chartOn && (
          <section className="display" aria-label="Gráfico">
            <div className="display-strip">
              <span>Gráfico · {rectLabel(rect)}</span>
              <button type="button" className="icon-btn" style={{ minHeight: 20 }} aria-label="Cerrar gráfico" onClick={() => setChartOn(false)}>
                ×
              </button>
            </div>
            <Chart values={values} rect={rect} />
          </section>
        )}

        <div className="tabs">
          {wb.sheets.map((s) =>
            renaming === s.id ? (
              <input
                key={s.id}
                className="tab-input"
                autoFocus
                defaultValue={s.name}
                aria-label="Nombre de la hoja"
                onBlur={(e) => {
                  edit({ type: 'renameSheet', sheetId: s.id, name: e.target.value });
                  setRenaming(null);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
                  if (e.key === 'Escape') setRenaming(null);
                }}
              />
            ) : (
              <button
                key={s.id}
                type="button"
                className="tab"
                aria-current={s.id === sheet.id ? 'page' : undefined}
                title="Doble clic para renombrar"
                onClick={() => setSheetId(s.id)}
                onDoubleClick={() => setRenaming(s.id)}
                onContextMenu={(e) => {
                  e.preventDefault();
                  if (wb.sheets.length > 1 && confirm(`¿Borrar la hoja "${s.name}"?`)) edit({ type: 'deleteSheet', sheetId: s.id });
                }}
              >
                {s.name}
              </button>
            ),
          )}
          <button type="button" className="tab" aria-label="Nueva hoja" onClick={() => edit({ type: 'addSheet' })}>
            +
          </button>
          <span className="status" role="status">
            {saveState === 'saving' ? 'Guardando…' : saveState === 'error' ? 'No se pudo guardar' : store.kind === 'cloud' ? 'Guardado en la nube' : 'Guardado en este navegador'}
          </span>
        </div>
      </main>
    </div>
  );
}
