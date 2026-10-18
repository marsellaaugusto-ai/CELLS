import type { Workbook, WorkbookSummary } from '../model/types';
import type { Book } from '../collab/book';

export interface StoredBook {
  /** Plain snapshot: what other apps read through the API. */
  wb: Workbook;
  /** Collaborative document (Yjs, base64). Missing for workbooks saved before collaboration. */
  state?: string;
  /** User id of the owner (cloud only). */
  owner?: string;
}

export interface Store {
  kind: 'local' | 'cloud';
  list(): Promise<WorkbookSummary[]>;
  load(id: string): Promise<StoredBook | null>;
  /** Saves the book; the cloud store first merges in whatever others saved meanwhile. */
  save(book: Book): Promise<void>;
  remove(id: string): Promise<void>;
}
