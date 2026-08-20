export interface ShadowCell {
  value: unknown;
  formula: string;
  numberFormat: string;
  /**
   * Column-level granularity, not a genuine per-cell snapshot (TASKS.md #64) —
   * see `agent.types.ts`'s `CellFormatCell` for the full explanation. Absent
   * means "not captured" (older context / no `formats` on the source SheetContext),
   * not "known to have no formatting".
   */
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}

export interface ShadowSheet {
  name: string;
  cells: Map<string, ShadowCell>;
  rowCount: number;
  columnCount: number;
  structure: string;
}

/**
 * One entry per structural (non-cell-shaped) mutation actually performed by virtualApply,
 * recorded with the *resolved* index/count it used — not the action's raw targeting syntax
 * (a column letter, a header name, a numeric index). diff.engine.ts's structural-op capture
 * reads this instead of re-parsing/re-resolving actions itself, so the two never drift apart.
 */
export interface StructuralLogEntry {
  opType: 'INSERT_COLUMN' | 'DELETE_COLUMN' | 'INSERT_ROW' | 'DELETE_ROW';
  sheetName: string;
  /**
   * For column ops: 0-based column index. For row ops: 1-based Excel row number
   * (matching `DeleteRowAction.rows`' own convention, so no conversion is needed
   * when building the inverse action).
   */
  index: number;
  count: number;
}

/**
 * Presence-only record of a conditional-format rule the shadow has seen
 * applied — NOT a simulation of the rule's actual per-cell fill effect (the
 * shadow has no fill/format state to evaluate against, same as every other
 * formatting-only action type). This exists so a verifier check can confirm
 * "a CONDITIONAL_FORMAT action for this range was actually applied" without
 * needing the shadow to reproduce Excel's own CF rendering (TASKS.md #39).
 */
export interface ShadowConditionalFormat {
  sheetName: string;
  range: string;
  ruleKind: string;
}

export interface ShadowWorkbook {
  activeSheetName: string;
  sheets: Map<string, ShadowSheet>;
  namedRanges: Map<string, string>;
  tables: string[];
  changedCells: Set<string>;
  structuralLog: StructuralLogEntry[];
  conditionalFormats: ShadowConditionalFormat[];
}
