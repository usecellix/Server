import {
  backoffDelayMs,
  classifyLlmError,
} from '../src/agents/utils/transient-llm-error.util';
import { AdaptiveConcurrency } from '../src/agents/utils/adaptive-concurrency.util';
import { LlmRequestError } from '../src/excel-ai/errors/llm-request.error';

/**
 * LONG_PROMPT_RELIABILITY_PLAN.md Phase 3 (TASKS.md #286).
 *
 * The failure this exists for, seen repeatedly across live runs:
 *   "OpenRouter could not verify available credits for this request in time.
 *    Retry shortly."
 * It arrives as a 402 — which `LlmRequestError.isRetryable` deliberately
 * excludes, on the sound reasoning that an empty wallet is still empty a
 * second later. But this one is the provider's credit CHECK timing out under
 * concurrency, and it killed whole month-sheet waves that had nothing wrong
 * with them.
 */
describe('classifyLlmError (TASKS.md #286)', () => {
  it('treats the live credit-check timeout as transient, despite being a 402', () => {
    const error = new LlmRequestError(
      402,
      'OpenRouter could not verify available credits for this request in time. Retry shortly.',
    );
    const { transient, reason } = classifyLlmError(error);
    expect(transient).toBe(true);
    expect(reason).toMatch(/credit-check timeout/i);
  });

  it('still treats a GENUINE insufficient-credit 402 as permanent — an empty wallet must fail fast', () => {
    const error = new LlmRequestError(402, 'Insufficient credits. Add more to continue.');
    expect(classifyLlmError(error)).toEqual({
      transient: false,
      reason: 'insufficient credit (402)',
    });
  });

  it('recognises the credit-check wording even when the status was lost in transport', () => {
    expect(classifyLlmError(new Error('could not verify available credits')).transient).toBe(true);
  });

  it('treats rate limits and server faults as transient', () => {
    for (const status of [429, 500, 502, 503, 504]) {
      expect(classifyLlmError(new LlmRequestError(status, 'x')).transient).toBe(true);
    }
  });

  it('treats auth and bad-request failures as permanent', () => {
    for (const status of [400, 401, 403, 404]) {
      expect(classifyLlmError(new LlmRequestError(status, 'x')).transient).toBe(false);
    }
  });

  it('treats transport faults as transient', () => {
    for (const message of ['socket hang up', 'ETIMEDOUT', 'ECONNRESET', 'request timed out']) {
      expect(classifyLlmError(new Error(message)).transient).toBe(true);
    }
  });

  it('reads a status off a plain object error too (SDK shapes vary)', () => {
    expect(classifyLlmError({ status: 429, message: 'slow down' }).transient).toBe(true);
    expect(classifyLlmError({ response: { status: 503 } }).transient).toBe(true);
  });

  it('never claims a decision it cannot explain', () => {
    expect(classifyLlmError(undefined).reason).toBeTruthy();
    expect(classifyLlmError(new Error('something odd')).reason).toBeTruthy();
  });
});

describe('backoffDelayMs', () => {
  it('grows exponentially and stays within the cap', () => {
    for (let attempt = 1; attempt <= 6; attempt += 1) {
      const delay = backoffDelayMs(attempt, 1000, 8000);
      expect(delay).toBeGreaterThanOrEqual(0);
      expect(delay).toBeLessThanOrEqual(8000);
    }
  });

  it('uses full jitter, so a throttled wave does not resynchronise on retry', () => {
    const samples = new Set(Array.from({ length: 40 }, () => backoffDelayMs(3)));
    expect(samples.size).toBeGreaterThan(1);
  });
});

describe('AdaptiveConcurrency (TASKS.md #286)', () => {
  it('starts where it is told, within floor and ceiling', () => {
    expect(new AdaptiveConcurrency({ start: 3 }).limit).toBe(3);
    expect(new AdaptiveConcurrency({ start: 99, ceiling: 6 }).limit).toBe(6);
    expect(new AdaptiveConcurrency({ start: 0, floor: 2 }).limit).toBe(2);
  });

  it('halves immediately when the provider pushes back', () => {
    const c = new AdaptiveConcurrency({ start: 6 });
    c.recordTransientFailure();
    expect(c.limit).toBe(3);
    c.recordTransientFailure();
    expect(c.limit).toBe(1);
  });

  it('never narrows below the floor', () => {
    const c = new AdaptiveConcurrency({ start: 2, floor: 1 });
    for (let i = 0; i < 10; i += 1) c.recordTransientFailure();
    expect(c.limit).toBe(1);
  });

  it('earns width back only after a RUN of successes, one slot at a time', () => {
    const c = new AdaptiveConcurrency({ start: 2, ceiling: 6, increaseAfterSuccesses: 4 });
    c.recordSuccess();
    c.recordSuccess();
    c.recordSuccess();
    expect(c.limit).toBe(2); // not yet
    c.recordSuccess();
    expect(c.limit).toBe(3); // one slot, not a jump back to the ceiling
  });

  it('never widens past the ceiling', () => {
    const c = new AdaptiveConcurrency({ start: 5, ceiling: 6, increaseAfterSuccesses: 1 });
    for (let i = 0; i < 20; i += 1) c.recordSuccess();
    expect(c.limit).toBe(6);
  });

  it('a transient fault resets the success run — width has to be re-earned from scratch', () => {
    const c = new AdaptiveConcurrency({ start: 4, increaseAfterSuccesses: 3 });
    c.recordSuccess();
    c.recordSuccess();
    c.recordTransientFailure(); // 4 -> 2, and the run of 2 successes is void
    expect(c.limit).toBe(2);
    c.recordSuccess();
    c.recordSuccess();
    expect(c.limit).toBe(2); // would have widened at 3 if the run had carried over
    c.recordSuccess();
    expect(c.limit).toBe(3);
  });
});
