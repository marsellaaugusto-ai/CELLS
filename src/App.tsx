import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react';
import { rectOf, rectLabel, toKey } from './model/address';
import { download, parseDelimited, toCsv } from './model/csv';
import { sampleWorkbook } from './model/sample';
import type { WorkbookSummary } from './model/types';
import { isError, newWorkbook } from './model/types';
import { evaluateWorkbook } from './engine/evaluate';
import { rectToTsv, usedRect, type Edit } from './state/workbook';
import type { ImportReport } from './io/xlsx';
import { parseCsv } from './model/csv';
import { Book, fromBase64 } from './collab/book';
import { makeMe, type Peer } from './collab/presence';
import { connect } from './collab/provider';
import { cloudTransport, tabTransport } from './collab/transport';
import { localStore } from './storage/local';
import { addMember, cloudStore, listMembers, removeMember, supabase } from './storage/cloud';
import type { Store, StoredBook } from './storage/store';
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
import { MODES, ModeHead, People, Rail, type ModeId } from './components/ui';

type SaveState = 'saved' | 'saving' | 'error';
export type CollabState = 'tabs' | 'cloud' | 'offline';

const PANELS = {
  archivo: ArchivoPanel,
  datos: DatosPanel,
  formato: FormatoPanel,
  formulas: FormulasPanel,
  graficos: GraficosPanel,
  conectar: ConectarPanel,
};

const bookFrom = (stored: StoredBook) => (stored.state ? new Book(fromBase64(stored.state)) : Book.fromWorkbook(stored.wb));

export default function App() {
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [userId, setUserId] = useState<string | null>(null);
  const [authChecked, setAuthChecked] = useState(!supabase);
  const store: Store = useMemo(() => (supabase && userEmail ? cloudStore(supabase) : localStore), [userEmail]);

  useEffect(() => {
    if (!supabase) return;
    supabase.auth.getSession().then(({ data }) => {
      setUserEmail(data.session?.user.email ?? null);
      setUserId(data.session?.user.id ?? null);
      setAuthChecked(true);
    });
    const { data } = supabase.auth.onAuthStateChange((_e, session) => {
      setUserEmail(session?.user.email ?? null);
      setUserId(session?.user.id ?? null);
    });
    return () => data.subscription.unsubscribe();
  }, []);

  const [book, setBook] = useState(() => Book.fromWorkbook(newWorkbook()));
  const wb = useSyncExternalStore(book.subscribe, () => book.snapshot);
  const edit = useCallback((e: Edit) => book.edit(e), [book]);
  const [owner, setOwner] = useState<string | undefined>();

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

  const refreshBooks = useCallback(() => store.list().then(setBooks).catch(() => setBooks([])), [store]);

  const open = useCallback(
    async (stored: StoredBook) => {
      const next = bookFrom(stored);
      // A workbook from before collaboration gets its shared document saved right away.
      if (!stored.state) await store.save(next).catch(() => setSaveState('error'));
      setBook(next);
      setOwner(stored.owner);
      setSheetId(next.snapshot.sheets[0].id);
      setSel({ anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } });
      setEditing(null);
    },
    [store],
  );

  // Boot (and re-boot when signing in/out): open the most recent workbook, or create the example.
  useEffect(() => {
    if (!authChecked) return;
    let cancelled = false;
    (async () => {
      const list = await store.list().catch(() => []);
      if (cancelled) return;
      const first = (list[0] && (await store.load(list[0].id).catch(() => null))) || { wb: sampleWorkbook() };
      if (cancelled) return;
      await open(first);
      setReady(true);
      refreshBooks();
    })();
    return () => {
      cancelled = true;
    };
  }, [store, authChecked, open, refreshBooks]);

  // Autosave shortly after each change made here. Changes from others are saved by whoever made them.
  useEffect(() => {
    if (!ready) return;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const save = () => {
      timer = undefined;
      store
        .save(book)
        .then(() => {
          setSaveState('saved');
          refreshBooks();
        })
        .catch(() => setSaveState('error'));
    };
    const onUpdate = (_u: Uint8Array, origin: unknown) => {
      if (!book.isLocal(origin)) return;
      setSaveState('saving');
      clearTimeout(timer);
      timer = setTimeout(save, 600);
    };
    book.doc.on('update', onUpdate);
    return () => {
      book.doc.off('update', onUpdate);
      if (timer) {
        clearTimeout(timer);
        save(); // don't lose the last edit when switching workbooks
      }
    };
  }, [book, ready, store, refreshBooks]);

  // Live collaboration: other tabs of this browser, or (with the cloud) everyone invited.
  const me = useMemo(() => makeMe(userEmail), [userEmail]);
  const [peers, setPeers] = useState<Peer[]>([]);
  const [collab, setCollab] = useState<CollabState>('tabs');
  const providerRef = useRef<ReturnType<typeof connect> | null>(null);

  useEffect(() => {
    if (!ready) return;
    const id = book.snapshot.id;
    const transport = store.kind === 'cloud' && supabase ? cloudTransport(supabase, id, me.id) : tabTransport(id);
    setCollab(store.kind === 'cloud' ? 'cloud' : 'tabs');
    transport.ready.catch(() => setCollab('offline'));
    const provider = connect(book, transport, me, setPeers);
    providerRef.current = provider;
    return () => {
      provider.close();
      providerRef.current = null;
      setPeers([]);
    };
  }, [book, ready, store, me]);

  useEffect(() => {
    providerRef.current?.moveCursor({ sheetId: sheet.id, r: sel.focus.r, c: sel.focus.c });
  }, [sheet.id, sel.focus.r, sel.focus.c, book, ready, me]);

  // Who can edit this workbook (cloud only; the owner manages the list).
  const isOwner = store.kind === 'cloud' && !!userId && owner === userId;
  const [members, setMembers] = useState<string[]>([]);
  useEffect(() => {
    setMembers([]);
    if (isOwner && supabase) listMembers(supabase, wb.id).then(setMembers).catch(() => setMembers([]));
  }, [isOwner, wb.id]);

  const evaluation = useMemo(() => evaluateWorkbook(wb), [wb, recalcTick]);
  const values = evaluation.values[sheet.id] ?? {};
  const approx = useMemo(() => {
    const prefix = `${sheet.id}!`;
    return new Set([...evaluation.approx].filter((k) => k.startsWith(prefix)).map((k) => k.slice(prefix.length)));
  }, [evaluation, sheet.id]);
  const [report, setReport] = useState<ImportReport | null>(null);
  const [busy, setBusy] = useState<string | null>(null);

  const openFile = async (file: File) => {
    setBusy(`Abriendo ${file.name}…`);
    try {
      if (/\.(xlsx|xlsm)$/i.test(file.name)) {
        const { readXlsx } = await import('./io/xlsx');
        const result = await readXlsx(await file.arrayBuffer(), file.name);
        await open({ wb: result.wb });
        setReport(result.report);
      } else {
        const next = newWorkbook(file.name.replace(/\.[^.]+$/, ''));
        const rows = parseCsv(await file.text());
        next.sheets[0].rows = Math.max(100, rows.length + 20);
        rows.forEach((row, r) => row.forEach((raw, c) => raw && (next.sheets[0].cells[toKey(r, c)] = { raw })));
        await open({ wb: next });
        setReport(null);
      }
      refreshBooks();
    } catch (e) {
      alert(`No se pudo abrir ${file.name}: ${e instanceof Error ? e.message : e}`);
    } finally {
      setBusy(null);
    }
  };

  const exportXlsx = async () => {
    setBusy('Preparando el .xlsx…');
    try {
      const { writeXlsx } = await import('./io/xlsx');
      const data = await writeXlsx(wb, evaluation.values);
      download(`${wb.name}.xlsx`, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } finally {
      setBusy(null);
    }
  };
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
    undo: () => book.undo(),
    redo: () => book.redo(),
    toggleBold: () => {
      const bold = !sheet.cells[activeKey]?.fmt?.bold;
      edit({ type: 'format', sheetId: sheet.id, rect, patch: { bold } });
    },
    addRows: () => edit({ type: 'addRows', sheetId: sheet.id, count: 50 }),
    resizeColumn: (c, px) =>
      edit({ type: 'layout', sheetId: sheet.id, patch: { colWidths: { ...sheet.layout?.colWidths, [c]: px } } }),
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
      if (next) await open(next);
    },
    newBook: async () => {
      await open({ wb: newWorkbook() });
      refreshBooks();
    },
    deleteBook: async (id) => {
      await store.remove(id).catch(() => setSaveState('error'));
      const list = await store.list().catch(() => []);
      setBooks(list);
      if (id === wb.id) {
        await open((list[0] && (await store.load(list[0].id).catch(() => null))) || { wb: newWorkbook() });
        refreshBooks();
      }
    },
    exportCsv,
    openFile,
    exportXlsx,
    report,
    clearReport: () => setReport(null),
    busy,
    exportJson: () => download(`${wb.name}.cells.json`, JSON.stringify(wb, null, 2), 'application/json'),
    storeKind: store.kind,
    cloudConfigured: !!supabase,
    userEmail,
    signIn,
    signOut: () => void supabase?.auth.signOut(),
    me,
    peers,
    collab,
    isOwner,
    members,
    invite: async (email) => {
      if (!supabase) return;
      await addMember(supabase, wb.id, email);
      setMembers(await listMembers(supabase, wb.id));
    },
    uninvite: async (email) => {
      if (!supabase) return;
      await removeMember(supabase, wb.id, email);
      setMembers(await listMembers(supabase, wb.id));
    },
  };

  const peerMarks = peers.flatMap((p) =>
    p.at && p.at.sheetId === sheet.id ? [{ key: toKey(p.at.r, p.at.c), name: p.name, color: p.color }] : [],
  );

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
            {ready ? <People me={me} peers={peers} /> : <span>Cargando…</span>}
          </div>
          <Grid
            sheet={sheet}
            values={values}
            sel={sel}
            editing={editing}
            show={show}
            actions={actions}
            approx={approx}
            peers={peerMarks}
            baseFont={wb.defaultFont}
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
