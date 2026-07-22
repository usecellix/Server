/** Base completion budget for a single-subtask verification. */
export const VERIFIER_BASE_MAX_TOKENS = 2000;

/** Extra tokens per subtask for per-subtask JSON results. */
export const VERIFIER_TOKENS_PER_SUBTASK = 180;

/** Cap so extreme chains don't blow the provider limit. */
export const VERIFIER_MAX_TOKENS_CAP = 8192;

/** Last-resort ceiling when the first verify call truncates / fails to parse. */
export const VERIFIER_LAST_RESORT_MAX_TOKENS = 8192;

export const VERIFIER_REASONING_MAX_TOKENS = 768;

/**
 * Scale verifier completion budget with subtask count so long chains
 * (dashboard / compound) have room for per-subtask JSON.
 */
export function resolveVerifierMaxTokens(subtaskCount: number): number {
  const n = Math.max(1, Math.floor(subtaskCount));
  return Math.min(VERIFIER_MAX_TOKENS_CAP, VERIFIER_BASE_MAX_TOKENS + n * VERIFIER_TOKENS_PER_SUBTASK);
}
