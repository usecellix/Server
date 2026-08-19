import { SheetActionType } from '../excel-ai/types/sheet-actions.types';

/**
 * Single source of truth for what virtualApply.ts's shadow-workbook simulator
 * actually simulates, for every action type — mirrors
 * excel-ai/types/action-catalog.ts's pattern (which does the same job for the
 * Executor's advertised-action list) applied to the Verifier's dry-run layer.
 *
 * Before this file, virtualApply.ts's switch had ~30 of 58 action types
 * falling through to an uncommented `break` or `default: break` with no
 * record of whether that silence was a deliberate "this type has no cell
 * effect" call or an oversight. Seven — DELETE_SHEET, FILL_DOWN, FILL_RIGHT,
 * CLEAR_CONTENT, CLEAR_ALL, SET_MATCHING_ROWS, MERGE_CELLS — turned out to
 * be genuine gaps: they mutate real values/structure but weren't simulated
 * at all, so the Verifier's dry-run couldn't see their effect before the
 * real Excel write. All are now simulated. See TASKS.md #41/#42/#66.
 */
type VirtualApplyCatalogEntry =
  /** Real simulation exists in virtualApply.ts's switch. */
  | { simulated: true }
  /** Deliberately not simulated — `reason` explains why. */
  | { simulated: false; reason: string };

const COSMETIC_VIEW_ONLY =
  'View/cosmetic only — no effect on cell values, formulas, or structure the shadow workbook diffs.';

export const VIRTUAL_APPLY_CATALOG: Record<SheetActionType, VirtualApplyCatalogEntry> = {
  // ---- Real simulations ----
  SET_CELL: { simulated: true },
  SET_FORMULA: { simulated: true },
  BATCH_SET: { simulated: true },
  WRITE_TABLE: { simulated: true },
  ADD_ROW: { simulated: true },
  DELETE_ROW: { simulated: true },
  INSERT_ROW: { simulated: true },
  INSERT_COLUMN: { simulated: true },
  DELETE_COLUMN: { simulated: true },
  SORT_RANGE: { simulated: true },
  MOVE_RANGE: { simulated: true },
  COPY_FILTERED_RANGE: { simulated: true },
  AGGREGATE_TABLE: { simulated: true },
  DEFINE_NAMED_RANGE: { simulated: true },
  ADD_SHEET: { simulated: true },
  CREATE_SHEET: { simulated: true },
  RENAME_SHEET: { simulated: true },
  COPY_SHEET: { simulated: true },
  DELETE_SHEET: { simulated: true },
  FILL_DOWN: { simulated: true },
  FILL_RIGHT: { simulated: true },
  CLEAR_CONTENT: { simulated: true },
  CLEAR_ALL: { simulated: true },
  SET_MATCHING_ROWS: { simulated: true },
  MERGE_CELLS: { simulated: true },

  // ---- Formatting-only, already documented in virtualApply.ts itself ----
  FORMAT_MATCHING_ROWS: {
    simulated: false,
    reason: 'Format-only — shadow workbook has no fill state to update.',
  },
  FORMAT_RANGE: {
    simulated: false,
    reason: 'Formatting-only (colours, borders, number-format display); no cell value change.',
  },
  HIGHLIGHT_CELL: { simulated: false, reason: 'Fill colour only; no cell value change.' },
  CLEAR_FORMAT: { simulated: false, reason: 'Formatting-only; no cell value change.' },

  // ---- Genuine view/cosmetic no-ops ----
  HIDE_ROW: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  UNHIDE_ROW: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SHOW_ROW: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  HIDE_COLUMN: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  UNHIDE_COLUMN: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SHOW_COLUMN: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SET_ROW_HEIGHT: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SET_COLUMN_WIDTH: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  AUTO_FILTER: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  FREEZE_PANES: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  UNFREEZE_PANES: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  AUTOFIT_COLUMNS: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  HIDE_SHEET: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SHOW_SHEET: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SET_SHEET_COLOR: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  PROTECT_SHEET: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  UNPROTECT_SHEET: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  SET_ZOOM: { simulated: false, reason: COSMETIC_VIEW_ONLY },
  ADD_COMMENT: { simulated: false, reason: 'Comments hang off the workbook, not cell values.' },
  DELETE_COMMENT: { simulated: false, reason: 'Comments hang off the workbook, not cell values.' },
  CREATE_TABLE: {
    simulated: false,
    reason: 'Wraps an existing range in an Excel Table structure; does not change cell values.',
  },
  CREATE_CHART: { simulated: false, reason: 'Charts are not cell data.' },
  UPDATE_CHART: { simulated: false, reason: 'Charts are not cell data.' },
  CLARIFY: { simulated: false, reason: 'Conversation control signal, not a sheet mutation.' },
  CHECKPOINT: { simulated: false, reason: 'Progress signal, not a sheet mutation.' },
  CLEAR_CELL: {
    simulated: false,
    reason:
      'Not advertised to the Executor (action-catalog.ts routes single-cell clears through CLEAR_CONTENT) — unreachable in practice today, flagged rather than silently assumed safe.',
  },

  // CLEAR_CONTENT, CLEAR_ALL, SET_MATCHING_ROWS, and MERGE_CELLS were found
  // as real gaps during #41's audit (they mutate values/structure but had no
  // simulation at all) and are now simulated — see the `simulated: true`
  // entries above. Fixed as TASKS.md #66.
  UNMERGE_CELLS: {
    simulated: false,
    reason: 'Reverses a merge; no cell-value effect of its own once MERGE_CELLS (#66) is simulated.',
  },
};

/** Every action type the backend catalog knows about, derived from the map above. */
export const ALL_VIRTUAL_APPLY_TYPES = Object.keys(
  VIRTUAL_APPLY_CATALOG,
) as SheetActionType[];

/** Action types the shadow workbook does not simulate, with their documented reason. */
export const UNSIMULATED_ACTION_TYPES = ALL_VIRTUAL_APPLY_TYPES.filter(
  (type) => !VIRTUAL_APPLY_CATALOG[type].simulated,
);
