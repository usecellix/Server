export class LlmRequestError extends Error {
  constructor(
    public readonly status: number,
    message?: string,
  ) {
    super(message ?? `LLM request failed (${status})`);
    this.name = 'LlmRequestError';
  }

  get isRecoverable(): boolean {
    return [402, 429, 500, 502, 503, 504].includes(this.status);
  }

  /**
   * Whether re-sending the SAME request could plausibly succeed.
   *
   * Distinct from `isRecoverable`, which only means "fall back rather than crash".
   * Task #92: a 402 is recoverable (degrade gracefully) but NOT retryable — an
   * empty wallet is still empty a second later, so retrying only doubles the wait
   * before the same failure. Transient server/network faults are retryable; a rate
   * limit is not retried inline because it needs a backoff we do not have here.
   */
  get isRetryable(): boolean {
    return [500, 502, 503, 504].includes(this.status);
  }
}

/** @deprecated use LlmRequestError */
export const OpenAiRequestError = LlmRequestError;
