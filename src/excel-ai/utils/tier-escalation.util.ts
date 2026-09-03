import { SheetActionPayload } from '../types/sheet-actions.types';

/**
 * Escalate a fast-lane result to the planner when the WORK turns out big —
 * TASKS.md #165.
 *
 * ## Why this exists
 *
 * `complexity-classifier.util.ts` decides which lane a prompt takes by reading
 * the prompt's *words*. That is a guess made before any work happens, and the
 * two failure directions are not symmetric:
 *
 *   - over-classified → the user waits longer than necessary (annoying)
 *   - under-classified → a multi-sheet build is answered by the single-action
 *     lane, with no planning, no consolidation and no verification (wrong)
 *
 * #165 fixed two guessing bugs (first-match-wins, and a narrow compound-phrase
 * list), but a word list can never be complete — TASKS.md #158 is this
 * codebase's own record of what leaning on one costs. So this is the net
 * underneath: after the fast lane has actually produced actions, look at what
 * it produced. The only reliable measure of a request's size is the work it
 * turned out to need.
 *
 * There is precedent in the classifier itself — `applyFindReplaceEscalation`
 * bumps lane 1 → 2 when a find/replace touches money. This generalises that
 * from "what words are in the prompt" to "what the request turned out to be".
 *
 * ## Why the thresholds are what they are
 *
 * Deliberately conservative, because escalating DISCARDS the fast lane's LLM
 * work and re-runs from the planner: the user pays that latency twice. A false
 * escalation costs seconds; a missed one costs their result. Both triggers
 * below are shapes the fast lanes are structurally unable to do well:
 *
 *   - **two or more sheet creations** — the fast lanes have no planner, so
 *     nothing decides what goes IN the second sheet, and the consolidation and
 *     presentation passes never see a sibling-schema group to work with.
 *   - **>= 25 actions** — the same floor `action-wave.util.ts` uses to decide a
 *     batch is big enough to stage. A single-action lane returning 25+ actions
 *     has already left its own design envelope.
 *
 * Both are counted on the fast lane's own emitted actions, so this makes no
 * assumption about vocabulary or domain (the #158 rule).
 */

/** Matches `MIN_ACTIONS_TO_STAGE` in `action-wave.util.ts` — same notion of "big". */
const ESCALATION_ACTION_THRESHOLD = 25;

/** Two created sheets means a structure, and a structure needs a plan. */
const ESCALATION_SHEET_CREATE_THRESHOLD = 2;

const SHEET_CREATING_TYPES = new Set(['ADD_SHEET', 'CREATE_SHEET', 'COPY_SHEET']);

export interface EscalationVerdict {
  escalate: boolean;
  /** Human-readable cause, logged and surfaced in telemetry. Null when not escalating. */
  reason: string | null;
}

export function assessTierEscalation(
  actions: SheetActionPayload[] | undefined | null,
): EscalationVerdict {
  if (!Array.isArray(actions) || actions.length === 0) {
    return { escalate: false, reason: null };
  }

  const sheetCreates = actions.filter((action) =>
    SHEET_CREATING_TYPES.has(String(action?.type)),
  ).length;

  if (sheetCreates >= ESCALATION_SHEET_CREATE_THRESHOLD) {
    return {
      escalate: true,
      reason: `fast lane emitted ${sheetCreates} sheet creations — a multi-sheet structure needs a plan`,
    };
  }

  if (actions.length >= ESCALATION_ACTION_THRESHOLD) {
    return {
      escalate: true,
      reason: `fast lane emitted ${actions.length} actions (>= ${ESCALATION_ACTION_THRESHOLD}) — beyond the single-action envelope`,
    };
  }

  return { escalate: false, reason: null };
}
