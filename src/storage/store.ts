import type { Workbook, WorkbookSummary } from '../model/types';

export interface Store {
  kind: 'local' | 'cloud';
  list(): Promise<WorkbookSummary[]>;
  load(id: string): Promise<Workbook | null>;
  save(wb: Workbook): Promise<void>;
  remove(id: string): Promise<void>;
  /** Calls back when someone else (another tab, device or app) saves this workbook. */
  subscribe(id: string, onChange: (wb: Workbook) => void): () => void;
}
