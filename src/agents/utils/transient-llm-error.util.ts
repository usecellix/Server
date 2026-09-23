/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 3 — telling a provider fault that will
 * pass apart from one that will not.
 *
 * The distinction matters because of one specific, repeatedly-observed live
 * failure: `OpenRouter could not verify available credits for this request in
 * time. Retry shortly.` It arrives as a 402, and `LlmRequestError.isRetryable`
 * deliberately excludes 402 on the sound reasoning that an empty wallet is
 * still empty a second later. But this particular 402 is not an empty wallet —
 * it is the provider's own credit check timing out under concurrency, and the
 * message says to retry. Across this session it failed whole month-sheet waves
 * that had nothing wrong with them.
 *
 * So the classification is deliberately narrow: a 402 is transient ONLY when
 * its message says the check could not be completed. A genuine
 * insufficient-credit 402 stays permanent and fails fast, exactly as before.
 */

/** A 402 that is a credit-CHECK timeout rather than an empty wallet. */
const CREDIT_CHECK_TIMEOUT =
  /could not verify (available )?credits|credit check (timed out|failed)|retry shortly/i;

/** Network/socket faults that never carry a usable HTTP status. */
const TRANSPORT_FAULT =
  /\b(etimedout|econnreset|econnrefused|epipe|enotfound|eai_again|socket hang up|network error|timeout|timed out|aborted)\b/i;

const TRANSIENT_STATUS = new Set([408, 409, 425, 429, 500, 502, 503, 504, 529]);

export interface TransientClassification {
  transient: boolean;
  /** Short, log-friendly reason — always present, so a decision is never unexplained. */
  reason: string;
}

function statusOf(error: unknown): number | null {
  if (!error || typeof error !== 'object') return null;
  const record = error as Record<string, unknown>;
  for (const key of ['status', 'statusCode', 'code']) {
    const value = record[key];
    if (typeof value === 'number' && value >= 100 && value < 600) return value;
  }
  const nested = record.response as Record<string, unknown> | undefined;
  if (nested && typeof nested.status === 'number') return nested.status;
  return null;
}

export function classifyLlmError(error: unknown): TransientClassification {
  const message = error instanceof Error ? error.message : String(error ?? '');
  const status = statusOf(error);

  if (status === 402) {
    return CREDIT_CHECK_TIMEOUT.test(message)
      ? { transient: true, reason: 'credit-check timeout (402, provider says retry)' }
      : { transient: false, reason: 'insufficient credit (402)' };
  }

  // The message can carry the credit-check wording even when the status was
  // lost in transport — the live failures reached us as a bare message.
  if (CREDIT_CHECK_TIMEOUT.test(message)) {
    return { transient: true, reason: 'credit-check timeout (by message)' };
  }

  if (status !== null && TRANSIENT_STATUS.has(status)) {
    return { transient: true, reason: `transient provider status ${status}` };
  }

  if (status !== null && status >= 400 && status < 500) {
    return { transient: false, reason: `client error ${status}` };
  }

  if (TRANSPORT_FAULT.test(message)) {
    return { transient: true, reason: 'transport fault' };
  }

  return { transient: false, reason: status ? `status ${status}` : 'unclassified' };
}

/** Exponential backoff with full jitter, so a throttled wave does not resynchronise. */
export function backoffDelayMs(attempt: number, baseMs = 1000, capMs = 8000): number {
  const ceiling = Math.min(capMs, baseMs * 2 ** Math.max(0, attempt - 1));
  return Math.round(Math.random() * ceiling);
}
