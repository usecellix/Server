import { LlmRequestError } from '../src/excel-ai/errors/llm-request.error';

/**
 * Task #92: the write path had no LLM retry at all — it fell back to the local rule
 * engine after 925ms, and that engine refuses structural writes, so one transient
 * provider fault meant total failure of the user's request.
 */
describe('LlmRequestError.isRetryable (#92)', () => {
  it('does not retry a 402 — an empty wallet is still empty a second later', () => {
    const err = new LlmRequestError(402, 'requires more credits');
    expect(err.isRecoverable).toBe(true); // still degrade gracefully
    expect(err.isRetryable).toBe(false); // but do not re-send
  });

  it('does not retry a rate limit inline (needs backoff we do not have here)', () => {
    expect(new LlmRequestError(429, 'Too Many Requests').isRetryable).toBe(false);
  });

  it.each([500, 502, 503, 504])('retries transient server fault %i', (status) => {
    expect(new LlmRequestError(status, 'upstream').isRetryable).toBe(true);
  });

  it('does not retry a client error', () => {
    expect(new LlmRequestError(400, 'bad request').isRetryable).toBe(false);
    expect(new LlmRequestError(401, 'unauthorized').isRetryable).toBe(false);
  });

  it('keeps isRecoverable and isRetryable independent', () => {
    // 402 is the case that motivated the split: recoverable but not retryable.
    const credits = new LlmRequestError(402, 'credits');
    expect(credits.isRecoverable && !credits.isRetryable).toBe(true);
    // A 400 is neither.
    const bad = new LlmRequestError(400, 'bad');
    expect(bad.isRecoverable || bad.isRetryable).toBe(false);
  });
});
