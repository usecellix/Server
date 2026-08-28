/**
 * Reasoning-token budgeting for reasoning models (gpt-5 family, o-series).
 *
 * For these models `maxCompletionTokens` is a SHARED budget: reasoning tokens and
 * content tokens both draw from it. With no reasoning cap the model can spend the
 * entire budget thinking and emit zero content — the request "succeeds", usage
 * shows completionTokens === reasoningTokens, and the caller sees an empty string.
 *
 * Real incident (2026-08-27): a read-only totals question on a ~14.5k-token sheet
 * payload returned empty twice at 4096 tokens, burning 85s, and was reported to the
 * user as a JSON parse failure. See TASKS.md F5.
 *
 * Fix: always reserve a floor for content by capping reasoning explicitly.
 */

/** Models whose completion budget is shared with reasoning tokens. */
const REASONING_MODEL_PATTERNS = [/(^|\/)gpt-5/i, /(^|\/)o[134](-|$)/i];

export function isReasoningModel(model: string | undefined): boolean {
  if (!model) return false;
  return REASONING_MODEL_PATTERNS.some((re) => re.test(model));
}

/** Fraction of the shared budget reasoning may consume. */
const REASONING_SHARE = 0.5;
/** Never cap reasoning below this — too small and the model degrades or errors. */
const MIN_REASONING_TOKENS = 512;
/** Content floor we try to preserve out of the shared budget. */
const MIN_CONTENT_TOKENS = 512;

/**
 * Reasoning cap for a shared completion budget, or undefined when no cap should be
 * sent (non-reasoning model, or a budget too small to split meaningfully — there the
 * caller must raise the budget instead; see `ensureBudgetForReasoning`).
 */
export function computeReasoningMaxTokens(
  model: string | undefined,
  completionBudget: number,
): number | undefined {
  if (!isReasoningModel(model)) return undefined;
  if (!Number.isFinite(completionBudget) || completionBudget <= 0) return undefined;
  if (completionBudget < MIN_REASONING_TOKENS + MIN_CONTENT_TOKENS) return undefined;

  const share = Math.floor(completionBudget * REASONING_SHARE);
  const capped = Math.min(share, completionBudget - MIN_CONTENT_TOKENS);
  return Math.max(MIN_REASONING_TOKENS, capped);
}

/**
 * A reasoning model needs room for BOTH phases. Callers tuned for non-reasoning
 * models pass budgets as low as 256, which cannot fit any reasoning at all — raise
 * the floor rather than starving the model.
 */
export function ensureBudgetForReasoning(
  model: string | undefined,
  requestedBudget: number,
): number {
  if (!isReasoningModel(model)) return requestedBudget;
  return Math.max(requestedBudget, MIN_REASONING_TOKENS + MIN_CONTENT_TOKENS);
}

/** Escalated budget for the retry after an empty (all-reasoning) completion. */
export function escalatedRetryBudget(budget: number): number {
  return Math.min(Math.max(budget * 2, 2048), 16384);
}
