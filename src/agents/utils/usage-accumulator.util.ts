import { LlmUsage } from '../../excel-ai/services/openrouter.service';

/**
 * Running token total across every real LLM call made within one orchestrator
 * run (Planner + every Executor/Verifier call across every subtask and retry).
 * Passed as a mutable out-param — the same pattern `LlmCompletionOutcome`
 * already uses on `OpenRouterService.complete()` — rather than returned,
 * because it needs to be threaded through and added to across many call sites
 * (Executor runs once per subtask per iteration; Verifier runs once per cycle)
 * that don't otherwise return a single combined value.
 */
export interface UsageTotals {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  /** Set once, by whichever caller resolves it first (conventionally the
   * Planner) — the model most representative of the run for audit purposes. */
  model?: string;
}

export function createUsageAccumulator(): UsageTotals {
  return { promptTokens: 0, completionTokens: 0, totalTokens: 0 };
}

/**
 * Add one LLM call's usage into the running total. Safe to call with
 * `undefined` (a call that never populated `outcome.usage`, e.g. because the
 * provider didn't return usage) — a no-op, not a NaN.
 */
export function addUsage(totals: UsageTotals | undefined, usage: LlmUsage | undefined): void {
  if (!totals || !usage) return;
  totals.promptTokens += usage.promptTokens ?? 0;
  totals.completionTokens += usage.completionTokens ?? 0;
  totals.totalTokens += usage.totalTokens ?? (usage.promptTokens ?? 0) + (usage.completionTokens ?? 0);
}
