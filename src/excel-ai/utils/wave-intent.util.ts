import { SheetActionPayload } from '../types/sheet-actions.types';

/**
 * Attribute finalized actions back to the plan subtask that produced them, so
 * each staged step can describe its OWN work — TASKS.md #167.
 *
 * ## The gap this closes
 *
 * TASKS.md #149 surfaced the Planner's own sentences ("Create sheet 'January'
 * and set A1:J1 headers") so the Accept card could say what a build DOES
 * instead of listing 40 mechanical actions. TASKS.md #161 then had to switch
 * that OFF for staged builds, because all six cards rendered the same
 * whole-plan bullets — "Create 13 sheets" was promising the formulas and charts
 * that arrive three steps later, which is #155's over-promising in a new shape.
 *
 * #161's remedy (fall back to the per-step action rollup) was correct and
 * deliberately temporary: the step label carried the semantics, so the bullets
 * could be mechanical without misleading anyone. But it left every staged card
 * describing machinery — "83 formatting changes" — when the information needed
 * to describe work was sitting one layer up the whole time.
 *
 * `OrchestratorRunResult.completedSubtasks` records exactly which actions each
 * subtask emitted. `finalizeActions` receives a FLAT array and the provenance
 * is dropped on the floor. That is `CODEBASE_ANALYSIS.md` §3.7's shape once
 * more — a signal the system had already computed and never passed to the code
 * that needed it.
 *
 * ## Why matching is structural, not referential
 *
 * The finalize pipeline rewrites what it is given: `sanitizeActions` merges
 * row-0 writes into a single ADD_ROW, the consolidation pass replaces formulas,
 * the presentation and chart passes append entirely new actions. So a finalized
 * action is frequently NOT the object the Executor emitted, and `===` matching
 * would attribute almost nothing.
 *
 * Matching therefore uses a coarse structural key (type + sheet + anchor).
 * Deliberately coarse: a false match between two same-shaped actions on the
 * same sheet costs nothing (both belong to that sheet's work anyway), whereas a
 * miss costs a bullet. Anything unmatched — every pass-generated format, freeze
 * and chart — falls back to sheet ownership, and only then to nothing.
 */

export interface SubtaskActions {
  subtaskId: string;
  actions: SheetActionPayload[];
}

export interface PlanIntentEntry {
  id: string;
  description: string;
  targetSheet: string;
}

/**
 * The sheet an action targets, whichever field carries it.
 *
 * `ADD_SHEET` is emitted as `{ type, name }` at least as often as
 * `{ type, sheetName }` — the live 140-action build opens with
 * `{"type":"ADD_SHEET","name":"Main"}`. Reading only `sheetName` silently
 * dropped that action from attribution entirely. Same concept under two field
 * names is ARCHITECTURE.md AD-7's drift pattern at the payload level, so it is
 * resolved in one place here rather than at each call site.
 */
export function resolveSheetName(action: SheetActionPayload): string {
  const candidates = [action?.sheetName, action?.name, action?.newSheetName];
  for (const candidate of candidates) {
    const text = typeof candidate === 'string' ? candidate.trim() : '';
    if (text) return text;
  }
  return '';
}

/**
 * A coarse identity for an action: enough to tell two different writes apart,
 * loose enough to survive the finalize passes' rewrites.
 */
export function structuralKey(action: SheetActionPayload): string {
  const type = String(action?.type ?? '');
  const sheet = resolveSheetName(action).toLowerCase();

  if (type === 'BATCH_SET' && Array.isArray(action.operations)) {
    const first = action.operations.find((op) => typeof op?.address === 'string');
    return `${type}|${sheet}|${String(first?.address ?? '').toUpperCase()}`;
  }

  const anchor =
    typeof action.address === 'string'
      ? action.address.toUpperCase()
      : typeof action.range === 'string'
        ? action.range.toUpperCase()
        : typeof action.row === 'number' && typeof action.col === 'number'
          ? `${action.row}:${action.col}`
          : '';

  return `${type}|${sheet}|${anchor}`;
}

/**
 * Which subtask each finalized action belongs to.
 *
 * Two passes: exact structural key first, then sheet ownership for everything
 * the passes invented. A sheet is "owned" by the subtask that created it, or
 * failing that by whichever subtask emitted the most actions on it — the
 * formatting for January belongs with January's work, whoever generated it.
 */
export function attributeActionsToSubtasks(
  finalized: SheetActionPayload[],
  completedSubtasks: SubtaskActions[],
): Map<number, string> {
  const byKey = new Map<string, string>();
  const sheetWeights = new Map<string, Map<string, number>>();

  for (const { subtaskId, actions } of completedSubtasks ?? []) {
    for (const action of actions ?? []) {
      const key = structuralKey(action);
      if (!byKey.has(key)) byKey.set(key, subtaskId);

      const sheet = resolveSheetName(action).toLowerCase();
      if (!sheet) continue;
      const weights = sheetWeights.get(sheet) ?? new Map<string, number>();
      weights.set(subtaskId, (weights.get(subtaskId) ?? 0) + 1);
      sheetWeights.set(sheet, weights);
    }
  }

  const sheetOwner = new Map<string, string>();
  for (const [sheet, weights] of sheetWeights) {
    let best: { id: string; count: number } | null = null;
    for (const [id, count] of weights) {
      if (!best || count > best.count) best = { id, count };
    }
    if (best) sheetOwner.set(sheet, best.id);
  }

  const attribution = new Map<number, string>();
  finalized.forEach((action, index) => {
    const exact = byKey.get(structuralKey(action));
    if (exact) {
      attribution.set(index, exact);
      return;
    }
    const sheet = resolveSheetName(action).toLowerCase();
    const owner = sheet ? sheetOwner.get(sheet) : undefined;
    if (owner) attribution.set(index, owner);
  });

  return attribution;
}

/**
 * The plan-intent entries a single wave's actions actually came from, in plan
 * order and de-duplicated.
 *
 * Returning `[]` is meaningful and expected: a wave built entirely from
 * pass-generated actions on sheets nothing owns has no intent to report, and
 * the caller must then fall back to #140's action rollup rather than render an
 * empty card — the failure TASKS.md #149 hit and fixed once already.
 */
export function intentForWave(
  waveActionIndexes: number[],
  attribution: Map<number, string>,
  planSubtasks: PlanIntentEntry[],
): PlanIntentEntry[] {
  const ids = new Set<string>();
  for (const index of waveActionIndexes) {
    const id = attribution.get(index);
    if (id) ids.add(id);
  }
  if (ids.size === 0) return [];
  return planSubtasks.filter((subtask) => ids.has(subtask.id));
}
