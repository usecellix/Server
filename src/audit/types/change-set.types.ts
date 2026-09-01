import { Action } from '../../agents/types/agent.types';
import {
  DomainException,
  SourceRef,
} from '../../domain-tools/types/domain-tool.types';

export type { DomainException, SourceRef };

export type ChangeSetStatus = 'previewed' | 'applied' | 'reverted';

export interface CellSnapshot {
  value: unknown;
  formula: string;
  format: string;
  /**
   * Column-level granularity, not a genuine per-cell snapshot (TASKS.md #64) —
   * see `agent.types.ts`'s `CellFormatCell` for the full explanation.
   */
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}

export interface CellChange {
  cell: string;
  sheet: string;
  before: unknown;
  after: unknown;
  formula?: string;
  isHardcoded: boolean;
  /** Optional citations — undefined for Tier 0/1 is expected */
  sourceRefs?: SourceRef[];
  /** Domain-tool exception flags — rendered distinctly in the UI */
  exceptionFlags?: DomainException[];
}

/**
 * One entry per structural action type as its inverse gets implemented (TASKS.md #12-17).
 * `opType` is typed as a union here for the engine's own exhaustiveness; the Mongoose
 * schema field itself stays a plain String so older/newer op types round-trip untyped.
 */
export type StructuralOpType =
  | 'ADD_SHEET'
  | 'DELETE_SHEET'
  | 'INSERT_COLUMN'
  | 'DELETE_COLUMN'
  | 'INSERT_ROW'
  | 'DELETE_ROW'
  | 'CREATE_TABLE'
  | 'MERGE_CELLS'
  | 'UNMERGE_CELLS'
  | 'CONDITIONAL_FORMAT'
  | 'CREATE_CHART';

export interface StructuralOp {
  opType: StructuralOpType;
  sheetName: string;
  params: Record<string, unknown>;
  appliedAt: Date;
}

/** PRD A6 signal (TASKS.md #49) — one Excel error string this change set introduced. */
export interface FormulaErrorChange {
  cell: string;
  sheet: string;
  error: string;
}

export interface ChangeSetRecord {
  changeSetId: string;
  conversationId: string;
  /** Durable per-workbook identity, resolved from the originating conversation (TASKS.md #24). */
  workbookId?: string;
  traceId: string;
  timestamp: Date;
  prompt: string;
  beforeState: Record<string, CellSnapshot>;
  changes: CellChange[];
  actions: Action[];
  /** Non-cell-shaped operations (sheet/column/row/chart/table/merge) — parallel to `changes`, not a replacement. */
  structuralOps: StructuralOp[];
  /**
   * Distinct action types in this change set with no defined revert path today
   * (TASKS.md #18) — empty when everything applied is fully revertible. Computed at
   * preview time so the caller can warn before Accept, not only discover it at revert.
   */
  irreversibleActionTypes: string[];
  /**
   * PRD A5 ("unintended modification rate", TASKS.md #48) — cell changes in this change
   * set on a sheet no action declared touching. Empty when every change fell within the
   * batch's own declared scope (the common case).
   */
  unintendedChanges: CellChange[];
  /**
   * PRD A6 ("formula error rate", TASKS.md #49) — Excel error strings this change set
   * introduced, distinct from any pre-existing error already sitting in the sheet.
   */
  formulaErrorsIntroduced: FormulaErrorChange[];
  status: ChangeSetStatus;
  appliedAt?: Date;
  revertedAt?: Date;
  /** Aggregate provenance for the change set (domain-tool confidence, etc.) */
  provenanceConfidence?: number;
}
