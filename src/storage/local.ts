import type { WorkbookSummary } from '../model/types';
import { toBase64 } from '../collab/book';
import type { Store, StoredBook } from './store';

const INDEX = 'cells:index';
const wbKey = (id: string) => `cells:wb:${id}`;
const stateKey = (id: string) => `cells:y:${id}`;

function readIndex(): WorkbookSummary[] {
  try {
    return JSON.parse(localStorage.getItem(INDEX) ?? '[]');
  } catch {
    return [];
  }
}

function writeIndex(list: WorkbookSummary[]) {
  localStorage.setItem(INDEX, JSON.stringify(list));
}

/** Saves in this browser only. Open tabs stay in sync through the tab transport. */
export const localStore: Store = {
  kind: 'local',

  async list() {
    return readIndex().sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  },

  async load(id) {
    const raw = localStorage.getItem(wbKey(id));
    if (!raw) return null;
    return { wb: JSON.parse(raw), state: localStorage.getItem(stateKey(id)) ?? undefined } satisfies StoredBook;
  },

  async save(book) {
    const wb = book.snapshot;
    localStorage.setItem(wbKey(wb.id), JSON.stringify(wb));
    localStorage.setItem(stateKey(wb.id), toBase64(book.encode()));
    const list = readIndex().filter((w) => w.id !== wb.id);
    list.push({ id: wb.id, name: wb.name, updatedAt: wb.updatedAt });
    writeIndex(list);
  },

  async remove(id) {
    localStorage.removeItem(wbKey(id));
    localStorage.removeItem(stateKey(id));
    writeIndex(readIndex().filter((w) => w.id !== id));
  },
};
