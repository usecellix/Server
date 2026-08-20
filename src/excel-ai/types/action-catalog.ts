import { SheetActionType } from './sheet-actions.types';

/**
 * Single source of truth for what the action catalog contains and what the
 * Executor is told about.
 *
 * Previously the catalog lived in three hand-maintained places — the
 * SheetActionType union, the normalizer's KNOWN_TYPES set, and a prose list in
 * the Executor system prompt. They drifted: FREEZE_PANES existed in the schema
 * and had a working handler, but was missing from the prompt, so the Tier 3
 * Executor could never emit it and "freeze the header row" silently did nothing.
 *
 * This map is typed Record<SheetActionType, …>, so adding a type to the union
 * without classifying it here fails the build. A capability can no longer exist
 * in the schema while staying invisible to the model by accident — omitting it
 * now has to be a decision someone wrote down.
 */
type CatalogEntry =
  /** Offered to the Executor in its available-action-types list. */
  | { advertise: true }
  /** Deliberately withheld — `reason` explains why, and is asserted by tests. */
  | { advertise: false; reason: string };

const ACTION_CATALOG: Record<SheetActionType, CatalogEntry> = {
  // ---- Cell / value writes ----
  SET_CELL: { advertise: true },
  SET_FORMULA: { advertise: true },
  HIGHLIGHT_CELL: { advertise: true },
  BATCH_SET: { advertise: true },
  WRITE_TABLE: { advertise: true },
  CLEAR_CELL: {
    advertise: false,
    reason: 'CLEAR_CONTENT covers single cells; a second near-identical verb invites the model to pick the wrong one.',
  },

  // ---- Rows / columns ----
  ADD_ROW: { advertise: true },
  DELETE_ROW: { advertise: true },
  INSERT_ROW: { advertise: true },
  INSERT_COLUMN: { advertise: true },
  DELETE_COLUMN: { advertise: true },
  HIDE_ROW: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  UNHIDE_ROW: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  SHOW_ROW: { advertise: false, reason: 'Alias of UNHIDE_ROW; Tier 0 lane resolves it.' },
  HIDE_COLUMN: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  UNHIDE_COLUMN: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  SHOW_COLUMN: { advertise: false, reason: 'Alias of UNHIDE_COLUMN; Tier 0 lane resolves it.' },
  SET_ROW_HEIGHT: { advertise: false, reason: 'Cosmetic sizing; Tier 0/1 lanes handle explicit requests.' },
  SET_COLUMN_WIDTH: {
    advertise: false,
    reason: 'AUTOFIT_COLUMNS is the advertised sizing action; explicit widths go through Tier 0/1.',
  },

  // ---- Range / layout ----
  FORMAT_RANGE: { advertise: true },
  FILL_DOWN: { advertise: true },
  FILL_RIGHT: { advertise: true },
  MERGE_CELLS: { advertise: true },
  SORT_RANGE: { advertise: true },
  MOVE_RANGE: { advertise: true },
  COPY_FILTERED_RANGE: { advertise: true },
  FORMAT_MATCHING_ROWS: { advertise: true },
  SET_MATCHING_ROWS: { advertise: true },
  CONDITIONAL_FORMAT: { advertise: true },
  DELETE_CONDITIONAL_FORMAT: {
    advertise: false,
    reason: 'Revert-only inverse of a CONDITIONAL_FORMAT create (TASKS.md #40) — not something the Executor should propose directly.',
  },
  AGGREGATE_TABLE: { advertise: true },
  AUTO_FILTER: { advertise: true },
  FREEZE_PANES: { advertise: true },
  UNFREEZE_PANES: { advertise: true },
  AUTOFIT_COLUMNS: { advertise: true },
  DEFINE_NAMED_RANGE: { advertise: true },
  CLEAR_CONTENT: { advertise: true },
  UNMERGE_CELLS: { advertise: false, reason: 'Rarely requested; MERGE_CELLS is the planned direction.' },
  CLEAR_FORMAT: { advertise: false, reason: 'FORMAT_RANGE with clearFill is the advertised way to strip formatting.' },
  CLEAR_ALL: {
    advertise: false,
    reason: 'Destructive and easy to over-apply; clear intents are routed deterministically instead.',
  },

  // ---- Sheets ----
  ADD_SHEET: { advertise: true },
  CREATE_SHEET: { advertise: true },
  DELETE_SHEET: { advertise: true },
  RENAME_SHEET: { advertise: true },
  COPY_SHEET: { advertise: true },
  HIDE_SHEET: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  SHOW_SHEET: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  SET_SHEET_COLOR: { advertise: false, reason: 'Cosmetic tab colour; Tier 0/1 lanes handle explicit requests.' },
  PROTECT_SHEET: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },
  UNPROTECT_SHEET: { advertise: false, reason: 'Handled deterministically by the Tier 0 shortcut lane.' },

  // ---- Objects ----
  CREATE_TABLE: { advertise: true },
  DELETE_TABLE: {
    advertise: false,
    reason: 'Revert-only inverse of CREATE_TABLE (TASKS.md #16) — not something the Executor should propose directly.',
  },
  CREATE_CHART: { advertise: true },
  UPDATE_CHART: { advertise: true },
  DELETE_CHART: {
    advertise: false,
    reason: 'Revert-only inverse of a CREATE_CHART create (TASKS.md #15) — not something the Executor should propose directly.',
  },
  ADD_COMMENT: { advertise: false, reason: 'Comments are not part of any current planned workflow.' },
  DELETE_COMMENT: { advertise: false, reason: 'Comments are not part of any current planned workflow.' },

  // ---- Not expressible in Office.js ----
  SET_ZOOM: {
    advertise: false,
    reason:
      'The Excel JavaScript API exposes no worksheet view zoom (only print zoom), so this action can never succeed.',
  },

  // ---- Control flow, not sheet mutations ----
  CLARIFY: { advertise: false, reason: 'Conversation control signal, not a sheet mutation.' },
  CHECKPOINT: { advertise: false, reason: 'Progress signal emitted by the loop, not a sheet mutation.' },
};

/** Every action type the system knows, derived from the exhaustive catalog. */
export const ALL_SHEET_ACTION_TYPES = Object.keys(ACTION_CATALOG) as SheetActionType[];

/** The action types offered to the Executor, in catalog order. */
export const EXECUTOR_ADVERTISED_ACTION_TYPES = ALL_SHEET_ACTION_TYPES.filter(
  (type) => ACTION_CATALOG[type].advertise,
);

/** Why a type is withheld from the Executor, or null when it is advertised. */
export function actionExclusionReason(type: SheetActionType): string | null {
  const entry = ACTION_CATALOG[type];
  return entry.advertise ? null : entry.reason;
}
