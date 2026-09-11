import { SheetActionPayload } from '../types/sheet-actions.types';
import { structuralKey } from './wave-intent.util';

/**
 * Decide which of a finished wave's actions are safe to show NOW — TASKS.md #174.
 *
 * ## The problem this solves
 *
 * A wave finishing is not the same as its actions being final. `finalizeActions`
 * runs three passes that need the WHOLE build in view:
 *
 *  - the consolidation pass REPLACES a Main formula once it can see every
 *    sibling sheet's schema,
 *  - the chart pass MOVES a chart anchor based on its source table's real width,
 *  - the presentation pass INVENTS header bands, number formats and freezes
 *    derived from header runs across all sheets.
 *
 * Emitting a card for an action one of those passes will later rewrite would be
 * a promise we then break: the user accepts "write this formula", and the
 * formula that eventually applies is a different one. That is exactly the
 * over-promising TASKS.md #155 and #161 were both about, and it is worse here
 * because the card would already be accepted.
 *
 * ## The rule
 *
 * Emit early ONLY the phases those passes never rewrite — sheet creation and
 * plain content writes. Everything else waits for the final pass, where it is
 * emitted exactly as today.
 *
 * That is not a compromise on the goal: creation and content are the bulk of a
 * large build's wall-clock time (12 month sheets, their headers and tables), so
 * showing them as they land is most of the perceived speed-up. The formula,
 * format, layout and chart phases arrive in the same cards they always did.
 *
 * ## Why the caller must de-duplicate
 *
 * The final emission still runs `finalizeActions` over EVERY action, because
 * the global passes need that. So an action shown early would otherwise appear
 * twice. `alreadyEmittedKeys` exists for that: the caller records what went out
 * and filters the final list. Matching is structural, not referential, for the
 * same reason `wave-intent.util.ts` documents — sanitize rewrites objects, so
 * `===` would match almost nothing.
 */

/** Action types the global finalize passes never rewrite. */
const EARLY_EMIT_TYPES = new Set<string>([
  // Creation — a sheet that exists is a fact no later pass revises.
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
  'RENAME_SHEET',
  // Plain content. NOTE: SET_FORMULA is deliberately absent — the consolidation
  // pass rewrites exactly those.
  'SET_CELL',
  'BATCH_SET',
  'ADD_ROW',
  'INSERT_ROW',
  'WRITE_TABLE',
]);

export function isEarlyEmittable(action: SheetActionPayload): boolean {
  return EARLY_EMIT_TYPES.has(String(action?.type));
}

/**
 * The subset of a wave that can be shown immediately, in the wave's own order.
 */
export function selectEarlyEmittable(
  waveActions: SheetActionPayload[],
): SheetActionPayload[] {
  return (waveActions ?? []).filter(isEarlyEmittable);
}

/** Types that bring a sheet into existence. */
const CREATE_TYPES = new Set(['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET', 'RENAME_SHEET']);

/**
 * Split a wave's early-emittable actions into CREATE-then-CONTENT groups.
 *
 * A single execution wave routinely contains both — "create sheet January and
 * write its headers" is one subtask — and emitting them as one card silently
 * discards the separation TASKS.md #160's phases were built on.
 *
 * The cost is concrete, and a live run showed it: a card labelled "Create and
 * fill 13 sheets" carried 12 ADD_SHEETs plus every header write. Because
 * `handleAddSheet` treats an existing sheet as a no-op while the overwrite
 * guard treats its occupied A1 as a refusal, running that card against a
 * workbook that already had those sheets blocked all 69 actions at once —
 * "target range A1 already contains data" — with no way to proceed. Split, the
 * creates apply harmlessly and only the genuinely-conflicting writes stop.
 *
 * Order matters and is preserved: creates first, then content, so a card can
 * never write into a sheet an unaccepted later card was going to make.
 * TASKS.md #175.
 */
export function splitEarlyByPhase(
  earlyActions: SheetActionPayload[],
): SheetActionPayload[][] {
  const creates = earlyActions.filter((a) => CREATE_TYPES.has(String(a?.type)));
  const content = earlyActions.filter((a) => !CREATE_TYPES.has(String(a?.type)));
  return [creates, content].filter((group) => group.length > 0);
}

/**
 * Remove actions already shown in a progressive card from the final list.
 *
 * Deliberately consumes each key at most once: if a build legitimately writes
 * two structurally identical actions, showing one early must not silently
 * swallow the other. A miss here costs a duplicated card row; over-filtering
 * would cost a real write.
 */
export function excludeAlreadyEmitted(
  finalActions: SheetActionPayload[],
  alreadyEmittedKeys: string[],
): SheetActionPayload[] {
  if (alreadyEmittedKeys.length === 0) return finalActions;

  const remaining = new Map<string, number>();
  for (const key of alreadyEmittedKeys) {
    remaining.set(key, (remaining.get(key) ?? 0) + 1);
  }

  return finalActions.filter((action) => {
    const key = structuralKey(action);
    const count = remaining.get(key);
    if (!count) return true;
    remaining.set(key, count - 1);
    return false;
  });
}

export function keysFor(actions: SheetActionPayload[]): string[] {
  return actions.map(structuralKey);
}
