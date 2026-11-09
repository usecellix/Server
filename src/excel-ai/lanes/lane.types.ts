import { SheetActionPayload } from '../types/sheet-actions.types';

/**
 * Deterministic lanes (Stage 2–4): a prompt the code can answer exactly is handled
 * here, without the LLM. A lane either answers (read-only) or proposes ordinary
 * actions that still go through preview → Accept → apply → revert.
 * See ACTION_REGISTRATION_CHECKLIST.md "First: do you need a new action type at all?".
 */
export interface LaneColumnMeta {
  header?: string;
  numberFormat?: string;
  formatCounts?: Array<{ format: string; count: number }>;
  formulaRows?: number[];
}

export interface LaneContext {
  message: string;
  /** Action mode — lanes that write are skipped otherwise. */
  writeAllowed: boolean;
  sheetName: string;
  values: unknown[][];
  headerRowIndex: number;
  columnMeta: Array<LaneColumnMeta | undefined>;
  sheetNames: string[];
  hiddenSheets: string[];
  /** Full values of another sheet (fetched on demand), or null if it does not exist. */
  getSheetValues: (name: string) => Promise<unknown[][] | null>;
  /** Excel 365 dynamic arrays available (client probe); undefined = unknown. */
  dynamicArrays?: boolean;
  now: Date;
}

export type LaneResult =
  | { kind: 'actions'; answer: string; explanation: string; actions: SheetActionPayload[] }
  | { kind: 'answer'; answer: string };

export interface DeterministicLane {
  /** Plan step it implements, e.g. "P2.5". */
  step: string;
  id: string;
  writes: boolean;
  detect: (message: string) => boolean;
  /** null = "not mine after all" → fall through to the normal (LLM) route. */
  run: (ctx: LaneContext) => LaneResult | null | Promise<LaneResult | null>;
}
