import type { Workbook, WorkbookSummary } from '../model/types';
import type { Store } from './store';

const INDEX = 'cells:index';
const wbKey = (id: string) => `cells:wb:${id}`;

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

/** Saves in this browser only. Other tabs of the same browser stay in sync. */
export const localStore: Store = {
  kind: 'local',

  async list() {
    return readIndex().sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  },

  async load(id) {
    const raw = localStorage.getItem(wbKey(id));
    return raw ? (JSON.parse(raw) as Workbook) : null;
  },

  async save(wb) {
    localStorage.setItem(wbKey(wb.id), JSON.stringify(wb));
    const list = readIndex().filter((w) => w.id !== wb.id);
    list.push({ id: wb.id, name: wb.name, updatedAt: wb.updatedAt });
    writeIndex(list);
  },

  async remove(id) {
    localStorage.removeItem(wbKey(id));
    writeIndex(readIndex().filter((w) => w.id !== id));
  },

  subscribe(id, onChange) {
    const handler = (e: StorageEvent) => {
      if (e.key === wbKey(id) && e.newValue) onChange(JSON.parse(e.newValue));
    };
    window.addEventListener('storage', handler);
    return () => window.removeEventListener('storage', handler);
  },
};
