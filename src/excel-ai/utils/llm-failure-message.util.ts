/**
 * Honest user-facing copy for a failed LLM call.
 *
 * F11 (2026-08-27): OpenRouter returned a precise, actionable 402 — "This request
 * requires more credits, or fewer max_tokens. You requested up to 4096 tokens, but
 * can only afford 2671." The pipeline discarded it and told the user to set
 * OPENROUTER_API_KEY — a key that was present and working. The product destroyed a
 * correct diagnosis and substituted a wrong one, sending the user to fix a
 * non-problem.
 *
 * Rule: never blame configuration for a runtime failure. If the provider told us
 * what went wrong, say that.
 */

export type LlmFailureKind =
  | 'not_configured'
  | 'out_of_credits'
  | 'rate_limited'
  | 'timeout'
  | 'provider_error'
  | 'unknown';

export type LlmFailure = {
  kind: LlmFailureKind;
  /** Raw provider detail, preserved for the user and the logs. */
  detail?: string;
  status?: number;
};

const CREDIT_PATTERNS = [/more credits/i, /can only afford/i, /insufficient (credit|balance|funds)/i];
const RATE_LIMIT_PATTERNS = [/rate limit/i, /too many requests/i];
const TIMEOUT_PATTERNS = [/timed? ?out/i, /etimedout/i, /econnreset/i, /socket hang up/i];

export function classifyLlmFailure(
  status: number | undefined,
  detail: string | undefined,
  isConfigured: boolean,
): LlmFailure {
  if (!isConfigured) {
    return { kind: 'not_configured', status, ...(detail ? { detail } : {}) };
  }
  const text = detail ?? '';
  if (status === 402 || CREDIT_PATTERNS.some((re) => re.test(text))) {
    return { kind: 'out_of_credits', status, ...(detail ? { detail } : {}) };
  }
  if (status === 429 || RATE_LIMIT_PATTERNS.some((re) => re.test(text))) {
    return { kind: 'rate_limited', status, ...(detail ? { detail } : {}) };
  }
  if (TIMEOUT_PATTERNS.some((re) => re.test(text))) {
    return { kind: 'timeout', status, ...(detail ? { detail } : {}) };
  }
  if (typeof status === 'number' && status >= 500) {
    return { kind: 'provider_error', status, ...(detail ? { detail } : {}) };
  }
  if (status !== undefined || detail) {
    return { kind: 'provider_error', status, ...(detail ? { detail } : {}) };
  }
  return { kind: 'unknown', status, ...(detail ? { detail } : {}) };
}

/** Copy for a failure that prevented a WRITE from being planned or applied. */
export function describeLlmFailureForWrite(failure: LlmFailure): string {
  const nothingApplied = '**Nothing was changed in your workbook.**';
  switch (failure.kind) {
    case 'not_configured':
      return (
        `I can't plan changes because no AI provider is configured. ${nothingApplied} ` +
        `Set OPENROUTER_API_KEY in the backend .env and restart the server.`
      );
    case 'out_of_credits':
      return (
        `Your AI provider account is out of credits, so I couldn't plan this change. ` +
        `${nothingApplied} Top up at https://openrouter.ai/settings/credits and send the ` +
        `request again.${failure.detail ? `\n\nProvider said: ${failure.detail}` : ''}`
      );
    case 'rate_limited':
      return (
        `The AI provider is rate-limiting requests right now, so I couldn't plan this ` +
        `change. ${nothingApplied} Wait a moment and try again.`
      );
    case 'timeout':
      return (
        `The AI provider didn't respond in time, so I couldn't plan this change. ` +
        `${nothingApplied} Try again — if it keeps happening, try a smaller request.`
      );
    default:
      return (
        `The AI provider failed, so I couldn't plan this change. ${nothingApplied} ` +
        `Try again.${failure.detail ? `\n\nProvider said: ${failure.detail}` : ''}`
      );
  }
}

/** Short status line streamed while falling back to the local engine. */
export function describeLlmFailureForStatus(failure: LlmFailure): string {
  switch (failure.kind) {
    case 'not_configured':
      return 'AI not configured — set OPENROUTER_API_KEY in backend .env';
    case 'out_of_credits':
      return 'AI provider out of credits — limited local mode…';
    case 'rate_limited':
      return 'AI provider rate-limited — limited local mode…';
    case 'timeout':
      return 'AI provider timed out — limited local mode…';
    default:
      return 'AI unavailable — limited local mode…';
  }
}
