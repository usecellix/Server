import { SheetActionType } from '../excel-ai/types/sheet-actions.types';

/**
 * Whether a real, working revert path exists TODAY for each action type — checked at
 * preview time (TASKS.md #18) so an unrevertable action is flagged before Accept, never
 * discovered only when the user tries to undo it.
 *
 * Deliberately NOT derived from `virtual-apply-catalog.ts`'s `simulated` flag, even though
 * the two look related. `simulated: true` only means "the shadow workbook's cell state
 * reflects this action's effect" — it does not mean a *correct* inverse exists. Three
 * types expose the gap directly:
 *   - RENAME_SHEET is `simulated: true`, but a rename isn't a cell-value change at all —
 *     the generic cell-diff machinery would misread the old sheet name's cells as entirely
 *     deleted and the new sheet name's cells as entirely new, producing a nonsense revert.
 *   - COPY_SHEET is `simulated: true`, but its "revert" (via the generic cell-level path)
 *     would only clear the copy's cell values — it never removes the sheet itself,
 *     leaving a phantom empty sheet behind. The same class of bug TASKS.md #13 fixed for
 *     column shifts, just not yet fixed here.
 *   - DEFINE_NAMED_RANGE is `simulated: true` (the shadow's `namedRanges` map is updated),
 *     but named-range bindings aren't `CellChange`-shaped at all, so nothing captures them
 *     for revert.
 * So this file is its own source of truth, checked against #10 (cell-level) and #12–17
 * (structural) directly, not inferred from a flag that answers a different question.
 */
type ReversibilityCatalogEntry =
  | { reversible: true }
  | { reversible: false; reason: string };

const COSMETIC_NOT_CAPTURED =
  'View/cosmetic action — not simulated by the shadow workbook, so nothing is captured to build a revert from.';

const NOT_A_MUTATION = 'Not a sheet mutation — never appears in an applied change set.';

export const REVERSIBILITY_CATALOG: Record<SheetActionType, ReversibilityCatalogEntry> = {
  // ---- Cell-level: generic before/after capture via diff.engine.ts's beforeStateToInverseActions
  // (TASKS.md #10). Safe whenever the action writes/clears cells in place without shifting or
  // removing the addressability of *other* cells — verified individually below, not assumed. ----
  SET_CELL: { reversible: true },
  SET_FORMULA: { reversible: true },
  BATCH_SET: { reversible: true },
  CLEAR_CONTENT: { reversible: true },
  CLEAR_ALL: { reversible: true },
  SET_MATCHING_ROWS: { reversible: true },
  SORT_RANGE: { reversible: true }, // permutes values within the same address set, no shift
  FILL_DOWN: { reversible: true },
  FILL_RIGHT: { reversible: true },
  MOVE_RANGE: { reversible: true }, // copies to dest + clears source; no shift of unrelated cells
  COPY_FILTERED_RANGE: { reversible: true }, // writes into a new/appended range
  AGGREGATE_TABLE: { reversible: true }, // writes new aggregate cells, append pattern
  WRITE_TABLE: { reversible: true },

  // ---- Structural: dedicated inverse-action construction in diff.engine.ts's
  // captureStructuralOps()/structuralOpsToInverseActions() ----
  ADD_SHEET: { reversible: true }, // TASKS.md #12
  CREATE_SHEET: { reversible: true }, // TASKS.md #12 (alias of ADD_SHEET)
  DELETE_SHEET: { reversible: true }, // TASKS.md #12
  INSERT_ROW: { reversible: true }, // TASKS.md #14
  DELETE_ROW: { reversible: true }, // TASKS.md #14
  ADD_ROW: { reversible: true }, // TASKS.md #14 — shares virtualAddRowAfter with INSERT_ROW
  INSERT_COLUMN: { reversible: true }, // TASKS.md #13
  DELETE_COLUMN: { reversible: true }, // TASKS.md #13
  CREATE_TABLE: { reversible: true }, // TASKS.md #16
  MERGE_CELLS: { reversible: true }, // TASKS.md #17
  UNMERGE_CELLS: { reversible: true }, // TASKS.md #17

  DELETE_TABLE: {
    reversible: false,
    reason:
      'Revert-only inverse of CREATE_TABLE (not advertised to the Executor — see action-catalog.ts). If ever applied as a forward action directly, nothing captures the original range/style needed to recreate the table.',
  },

  // ---- Simulated but genuinely NOT revertible today — the gap this catalog exists to catch ----
  RENAME_SHEET: {
    reversible: false,
    reason:
      "A rename isn't a cell-value change — the generic cell-diff would treat the old sheet name's cells as deleted and the new name's cells as newly created, producing an incorrect revert rather than renaming back.",
  },
  COPY_SHEET: {
    reversible: false,
    reason:
      "The generic cell-level revert would clear the copy's values but never removes the sheet itself, leaving a phantom empty sheet — no structural inverse has been built for this yet.",
  },
  DEFINE_NAMED_RANGE: {
    reversible: false,
    reason: "Named-range bindings aren't cell data — nothing captures the previous binding (or its absence) to restore.",
  },

  // ---- Formatting-only — not simulated, so no before/after state is ever captured ----
  FORMAT_RANGE: { reversible: false, reason: 'Formatting-only; not simulated, so no format state is captured to restore.' },
  FORMAT_MATCHING_ROWS: { reversible: false, reason: 'Formatting-only; not simulated, so no format state is captured to restore.' },
  HIGHLIGHT_CELL: { reversible: false, reason: 'Fill colour only; not simulated, so no prior fill state is captured.' },
  CLEAR_FORMAT: { reversible: false, reason: 'Formatting-only; not simulated, so no format state is captured to restore.' },
  CLEAR_CELL: {
    reversible: false,
    reason: 'Not advertised to the Executor and not simulated by virtualApply.ts — unreachable in practice, but flagged rather than silently assumed safe.',
  },

  // ---- Runtime-assigned-id problem (TASKS.md #15) — solved using the same apply-endpoint
  // contract mechanism TASKS.md #40 built first for CONDITIONAL_FORMAT: the frontend reads
  // back the real Excel-assigned chart name right after creating it, reports it to
  // POST /audit/apply/:changeSetId, and ChangeSetService.markApplied() patches it into the
  // pending CREATE_CHART structuralOp before the change set is marked applied. ----
  CREATE_CHART: { reversible: true },
  DELETE_CHART: {
    reversible: false,
    reason:
      'Revert-only inverse of a CREATE_CHART create (TASKS.md #15) — not advertised to the Executor. If ever applied as a forward action directly, nothing captures the deleted chart\'s original configuration needed to recreate it.',
  },
  // CONDITIONAL_FORMAT's base-case entry is `reversible: true` (TASKS.md #40) — a create
  // now has a real inverse (DELETE_CONDITIONAL_FORMAT, targeting the real Excel-assigned
  // rule id captured at apply time and patched into structuralOps before the inverse is
  // built). But an action instance with `existingRuleId` set is a MODIFY of an existing
  // rule in place — the rule's prior parameters aren't captured anywhere (context only
  // keeps a human-readable `summary`, not the full rule object), so that specific instance
  // stays irreversible. This per-instance nuance can't be expressed by this catalog's
  // flat per-type shape, so `computeIrreversibleActionTypes()` below special-cases it
  // directly rather than reading this entry for CONDITIONAL_FORMAT at all.
  CONDITIONAL_FORMAT: { reversible: true },
  DELETE_CONDITIONAL_FORMAT: {
    reversible: false,
    reason:
      'Revert-only inverse of a CONDITIONAL_FORMAT create (TASKS.md #40) — not advertised to the Executor. If ever applied as a forward action directly, nothing captures the deleted rule\'s original parameters needed to recreate it.',
  },
  UPDATE_CHART: {
    reversible: false,
    reason: "Modifies an existing chart's type/colour scheme; nothing captures the chart's prior configuration to restore.",
  },

  // ---- Genuine view/cosmetic no-ops — never simulated, nothing to capture ----
  HIDE_ROW: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  UNHIDE_ROW: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SHOW_ROW: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  HIDE_COLUMN: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  UNHIDE_COLUMN: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SHOW_COLUMN: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SET_ROW_HEIGHT: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SET_COLUMN_WIDTH: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  FREEZE_PANES: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  UNFREEZE_PANES: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  AUTO_FILTER: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  AUTOFIT_COLUMNS: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SET_ZOOM: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  PROTECT_SHEET: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  UNPROTECT_SHEET: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  HIDE_SHEET: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SHOW_SHEET: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  SET_SHEET_COLOR: { reversible: false, reason: COSMETIC_NOT_CAPTURED },
  ADD_COMMENT: { reversible: false, reason: 'Comments hang off the workbook, not cell values — not simulated, nothing captured.' },
  DELETE_COMMENT: { reversible: false, reason: 'Comments hang off the workbook, not cell values — not simulated, nothing captured.' },

  // ---- Control signals — not mutations at all ----
  CLARIFY: { reversible: false, reason: NOT_A_MUTATION },
  CHECKPOINT: { reversible: false, reason: NOT_A_MUTATION },
};

/** Every action type this catalog classifies, derived from the map above. */
export const ALL_REVERSIBILITY_CATALOG_TYPES = Object.keys(
  REVERSIBILITY_CATALOG,
) as SheetActionType[];

/**
 * Accepts either the plain type strings this function originally took, or the fuller
 * action-shaped objects `change-set.service.ts` actually has in hand — the latter is what
 * lets the CONDITIONAL_FORMAT special case below see `existingRuleId`, which the catalog's
 * flat per-type shape can't express on its own.
 */
type IrreversibilityCheckInput = string | { type: string; existingRuleId?: string };

/**
 * Given the action types (or action instances) present in a change set, return the distinct
 * types that have no defined revert path today — empty if every action is fully revertible.
 * Unknown/unlisted types are treated as irreversible (fail closed) rather than silently
 * assumed safe.
 *
 * CONDITIONAL_FORMAT is a per-instance exception, not read from the catalog directly: a
 * plain create (no `existingRuleId`) is reversible per TASKS.md #40, but an instance that
 * sets `existingRuleId` modifies an existing rule in place, and nothing captures that rule's
 * prior parameters to restore — so that specific instance is flagged irreversible even
 * though the catalog's own `CONDITIONAL_FORMAT` entry says `reversible: true`.
 */
export function computeIrreversibleActionTypes(
  actionTypes: IrreversibilityCheckInput[],
): string[] {
  const irreversible = new Set<string>();
  for (const entry of actionTypes) {
    const type = typeof entry === 'string' ? entry : entry.type;
    const existingRuleId = typeof entry === 'string' ? undefined : entry.existingRuleId;

    if (type === 'CONDITIONAL_FORMAT' && existingRuleId) {
      irreversible.add(type);
      continue;
    }

    const catalogEntry = REVERSIBILITY_CATALOG[type as SheetActionType];
    if (!catalogEntry || !catalogEntry.reversible) {
      irreversible.add(type);
    }
  }
  return [...irreversible];
}
