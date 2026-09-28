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
const REASONING_MODEL_PATTERNS = [/(^|\/)gpt-5/i, /(^|\/)o[134](-|$)/i, /(^|\/)glm-/i];

export function isReasoningModel(model: string | undefined): boolean {
  if (!model) return false;
  return REASONING_MODEL_PATTERNS.some((re) => re.test(model));
}

/**
 * Models that reject `reasoning.effort: 'none'` outright (OpenRouter 400:
 * "Reasoning is mandatory for this endpoint and cannot be disabled") rather
 * than merely sharing their completion budget with reasoning tokens. This is
 * a stricter subset of `isReasoningModel` in principle, but in practice every
 * GLM model this deployment has actually sent a request to rejects `'none'`
 * 100% of the time (confirmed directly against `llm_calls`: every attempt-1
 * call with `effort: 'none'` to `z-ai/glm-5.3` or `z-ai/glm-5.3-flash` failed
 * with this exact error, every attempt-2 retry at `effort: 'low'` succeeded)
 * — so GLM is listed here too rather than assumed compatible with `'none'`.
 *
 * `OpenRouterService.requestChatCompletion` already retries a rejected
 * `'none'` request once at `effort: 'low'` (TASKS.md's reasoning-mandatory
 * fix), which is what keeps every affected call from actually failing — but
 * for a model on THIS list, that first attempt is not a real attempt, it is
 * a guaranteed-failing round trip paid on every single request through that
 * lane (TASKS.md #228). `minReasoningEffort` lets a caller skip straight to
 * the effort level that actually works.
 */
const REASONING_MANDATORY_MODEL_PATTERNS = [/(^|\/)glm-/i];

export function isReasoningMandatory(model: string | undefined): boolean {
  if (!model) return false;
  return REASONING_MANDATORY_MODEL_PATTERNS.some((re) => re.test(model));
}

/**
 * The lowest `reasoning.effort` value that will not be rejected outright for
 * `model`. Callers that would otherwise default to `'none'` should use this
 * instead — it degrades gracefully to `'none'` for any model not known to
 * require reasoning, so it's always safe to route a fixed `'none'` request
 * through this rather than special-casing each caller.
 */
export function minReasoningEffort(model: string | undefined): 'none' | 'low' {
  return isReasoningMandatory(model) ? 'low' : 'none';
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
