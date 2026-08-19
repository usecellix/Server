/**
 * The Executor's completion budget used to be a flat 2000 tokens regardless of
 * how much a subtask actually asked for — unlike the Planner and Verifier,
 * which already scale (see planner-token-budget.util.ts, spec 16; and
 * verifier-token-budget.util.ts). A subtask writing a formula table across many
 * sheets (e.g. one row per month, several SUMIF formulas per row) can need far
 * more than 2000 tokens of JSON output. When it truncates, the retry re-runs
 * the SAME subtask with the SAME fixed budget and fails the same way again —
 * which is exactly what "failed after 2 attempts, formulas only through April"
 * looked like: both attempts hit the same ceiling.
 */

/** Base completion budget for a small subtask (a handful of writes). */
export const EXECUTOR_BASE_MAX_TOKENS = 2000;

/** Extra tokens per estimated action — each is roughly one JSON action object. */
export const EXECUTOR_TOKENS_PER_ACTION = 60;

/** Cap so a runaway estimate doesn't blow the provider limit. */
export const EXECUTOR_MAX_TOKENS_CAP = 8192;

/**
 * Scale the Executor's completion budget with the subtask's own estimated
 * action count, so a large table of formulas gets room to finish in one call
 * instead of truncating and failing retry after retry against the same cap.
 */
export function resolveExecutorMaxTokens(estimatedActions?: number): number {
  const n = Math.max(1, Math.floor(estimatedActions ?? 1));
  return Math.min(
    EXECUTOR_MAX_TOKENS_CAP,
    EXECUTOR_BASE_MAX_TOKENS + n * EXECUTOR_TOKENS_PER_ACTION,
  );
}
