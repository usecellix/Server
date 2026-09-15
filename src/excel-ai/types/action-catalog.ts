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
  // The Tier 0 lane only resolves these when the user names a column LETTER.
  // "Hide the Narration column" (a header name) escalates to Tier 3, where the
  // Executor had no hide verb and improvised SET_COLUMN_WIDTH width:0 with a
  // `columns: ["I"]` field that sanitizeAction then dropped — "1 action,
  // verified: true" followed by "Something went wrong". Third instance of the
  // #166 pattern. TASKS.md #215.
  HIDE_ROW: { advertise: true },
  UNHIDE_ROW: { advertise: true },
  SHOW_ROW: { advertise: false, reason: 'Alias of UNHIDE_ROW; UNHIDE_ROW is the advertised spelling.' },
  HIDE_COLUMN: { advertise: true },
  UNHIDE_COLUMN: { advertise: true },
  SHOW_COLUMN: { advertise: false, reason: 'Alias of UNHIDE_COLUMN; UNHIDE_COLUMN is the advertised spelling.' },
  // Fourth instance of the #166 pattern: the Tier 0 lane only matches the
  // literal "set row height" wording, so "Set the height of rows 2 to 5 to 25"
  // reached Tier 3, where the Executor had the type name but no schema and
  // emitted a bare { type, sheetName } that sanitizeAction dropped. Advertised
  // with a schema alongside SET_COLUMN_WIDTH (#169). TASKS.md #216.
  SET_ROW_HEIGHT: { advertise: true },
  // Was withheld because "explicit widths go through Tier 0/1" — the same false
  // premise HIDE_SHEET carried (#166): those lanes never run inside a Tier 3
  // build, so a dashboard that wants a wide label column and narrow number
  // columns had no way to say so and got autofit-to-content everywhere.
  // TASKS.md #169.
  SET_COLUMN_WIDTH: { advertise: true },

  // ---- Range / layout ----
  FORMAT_RANGE: { advertise: true },
  FILL_DOWN: { advertise: true },
  FILL_RIGHT: { advertise: true },
  MERGE_CELLS: { advertise: true },
  SORT_RANGE: { advertise: true },
  SET_RANGE_VALUES: {
    advertise: false,
    reason: 'Revert-only bulk inverse (TASKS.md #100, fast-path for large reverts like undoing a SORT_RANGE) — not something the Executor should propose directly.',
  },
  MOVE_RANGE: { advertise: true },
  COPY_FILTERED_RANGE: { advertise: true },
  FORMAT_MATCHING_ROWS: { advertise: true },
  SET_MATCHING_ROWS: { advertise: true },
  // The only safe way to answer "delete blank rows" / "delete rows where X":
  // which rows match is resolved against the real cells at apply time instead
  // of being guessed as a DELETE_ROW row/rowCount pair. TASKS.md #234.
  DELETE_MATCHING_ROWS: { advertise: true },
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
  HIDE_GRIDLINES: { advertise: true },
  DATA_VALIDATION: { advertise: true },
  DEFINE_NAMED_RANGE: { advertise: true },
  CLEAR_CONTENT: { advertise: true },
  // "Unmerge all merged cells in this sheet" and "Clear all formatting in
  // A1:I31" are both guide use cases (T1.3) that reach Tier 3; withholding the
  // verbs made the Executor improvise shapes that were dropped. TASKS.md #215.
  UNMERGE_CELLS: { advertise: true },
  CLEAR_FORMAT: { advertise: true },
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
  // Guide T1.1's "move sheet" had no action at all, so the Executor built it
  // out of copy + rename + delete — a plan that passed deterministic checks
  // and would have destroyed the sheet. TASKS.md #212.
  MOVE_SHEET: { advertise: true },
  // Was withheld as "Tier 0 handles it" — true for *"hide the Lists sheet"* as a
  // standalone request, and false for the case that matters: a Tier 3 build that
  // creates a lookup sheet to back its dropdowns and wants it out of the way. The
  // Tier 0 lane never runs inside a Tier 3 plan, so the capability was
  // unreachable exactly when it was needed. This is the FREEZE_PANES bug this
  // file's own header describes, in a second form. TASKS.md #166.
  HIDE_SHEET: { advertise: true },
  // Withheld as "Tier 0 handles it" — the same false premise as HIDE_SHEET
  // (#166) and SET_COLUMN_WIDTH (#169), and worse here: HIDE_SHEET *was*
  // advertised, so "Unhide the Working sheet" reached the Executor with hide as
  // the only sheet-visibility verb it had been given and it hid the sheet —
  // the exact opposite of the request. TASKS.md #211.
  SHOW_SHEET: { advertise: true },
  // Same trap: "Change the Summary tab colour to blue" is not matched by the
  // Tier 0 regex (it colours the active sheet only), so it escalated to Tier 3,
  // where the Executor had no way to express it and burned two retries before
  // failing with "could not complete and verify". TASKS.md #211.
  SET_SHEET_COLOR: { advertise: true },
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
  // Guide T1.3 lists "add / remove comment" as a Tier 1 operation, and the
  // router already routes it here. TASKS.md #215.
  ADD_COMMENT: { advertise: true },
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
