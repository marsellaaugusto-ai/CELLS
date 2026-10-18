import type { WorkbookSummary } from '../model/types';
import { toBase64 } from '../collab/book';
import type { Store, StoredBook } from './store';

// Saves in this browser with IndexedDB, which holds hundreds of MB; localStorage (used before)
// stops at about 5 MB, too little for real Excel files. Open tabs stay in sync through the
// tab transport.

const DB = 'cells';
const BOOKS = 'books'; // id → StoredBook
const INDEX = 'index'; // id → WorkbookSummary (small, so listing doesn't load every book)

let dbPromise: Promise<IDBDatabase> | null = null;

function db(): Promise<IDBDatabase> {
  dbPromise ??= new Promise<IDBDatabase>((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(BOOKS);
      req.result.createObjectStore(INDEX);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }).then(async (d) => {
    await migrateFromLocalStorage(d);
    return d;
  });
  return dbPromise!;
}

function run<T>(d: IDBDatabase, stores: string[], mode: IDBTransactionMode, fn: (tx: IDBTransaction) => IDBRequest<T> | void) {
  return new Promise<T>((resolve, reject) => {
    const tx = d.transaction(stores, mode);
    const req = fn(tx);
    tx.oncomplete = () => resolve(req ? req.result : (undefined as T));
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

/** Books saved by earlier versions in localStorage move to IndexedDB once. */
async function migrateFromLocalStorage(d: IDBDatabase) {
  let index: WorkbookSummary[];
  try {
    index = JSON.parse(localStorage.getItem('cells:index') ?? '[]');
  } catch {
    return;
  }
  if (!index.length) return;
  await run(d, [BOOKS, INDEX], 'readwrite', (tx) => {
    for (const summary of index) {
      const raw = localStorage.getItem(`cells:wb:${summary.id}`);
      if (!raw) continue;
      const book: StoredBook = { wb: JSON.parse(raw), state: localStorage.getItem(`cells:y:${summary.id}`) ?? undefined };
      tx.objectStore(BOOKS).put(book, summary.id);
      tx.objectStore(INDEX).put(summary, summary.id);
    }
  });
  for (const summary of index) {
    localStorage.removeItem(`cells:wb:${summary.id}`);
    localStorage.removeItem(`cells:y:${summary.id}`);
  }
  localStorage.removeItem('cells:index');
}

export const localStore: Store = {
  kind: 'local',

  async list() {
    const d = await db();
    const all = await run<WorkbookSummary[]>(d, [INDEX], 'readonly', (tx) => tx.objectStore(INDEX).getAll());
    return all.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  },

  async load(id) {
    const d = await db();
    return (await run<StoredBook | undefined>(d, [BOOKS], 'readonly', (tx) => tx.objectStore(BOOKS).get(id))) ?? null;
  },

  async save(book) {
    const wb = book.snapshot;
    const stored: StoredBook = { wb, state: toBase64(book.encode()) };
    const d = await db();
    await run(d, [BOOKS, INDEX], 'readwrite', (tx) => {
      tx.objectStore(BOOKS).put(stored, wb.id);
      tx.objectStore(INDEX).put({ id: wb.id, name: wb.name, updatedAt: wb.updatedAt } satisfies WorkbookSummary, wb.id);
    });
  },

  async remove(id) {
    const d = await db();
    await run(d, [BOOKS, INDEX], 'readwrite', (tx) => {
      tx.objectStore(BOOKS).delete(id);
      tx.objectStore(INDEX).delete(id);
    });
  },
};
