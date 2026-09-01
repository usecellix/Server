import {
  classifyLlmFailure,
  describeLlmFailureForStatus,
  describeLlmFailureForWrite,
} from '../src/excel-ai/utils/llm-failure-message.util';

/** The exact text OpenRouter returned in the F11 incident. */
const REAL_402 =
  'AI provider failed (402): This request requires more credits, or fewer max_tokens. ' +
  'You requested up to 4096 tokens, but can only afford 2671.';

describe('classifyLlmFailure', () => {
  it('classifies the real 402 as out_of_credits, not a config problem', () => {
    expect(classifyLlmFailure(402, REAL_402, true).kind).toBe('out_of_credits');
  });

  it('detects credit exhaustion from the message even without a 402 status', () => {
    expect(classifyLlmFailure(undefined, 'can only afford 2671', true).kind).toBe(
      'out_of_credits',
    );
  });

  it('only reports not_configured when the key is genuinely absent', () => {
    expect(classifyLlmFailure(undefined, undefined, false).kind).toBe('not_configured');
    // Key present + provider failure must NEVER be reported as not_configured.
    expect(classifyLlmFailure(402, REAL_402, true).kind).not.toBe('not_configured');
    expect(classifyLlmFailure(429, 'rate limit', true).kind).not.toBe('not_configured');
    expect(classifyLlmFailure(503, 'unavailable', true).kind).not.toBe('not_configured');
  });

  it('classifies rate limits and timeouts distinctly', () => {
    expect(classifyLlmFailure(429, 'Too Many Requests', true).kind).toBe('rate_limited');
    expect(classifyLlmFailure(undefined, 'socket hang up', true).kind).toBe('timeout');
    expect(classifyLlmFailure(undefined, 'ETIMEDOUT', true).kind).toBe('timeout');
  });

  it('classifies 5xx as a provider error', () => {
    expect(classifyLlmFailure(503, 'Service Unavailable', true).kind).toBe('provider_error');
  });
});

describe('describeLlmFailureForWrite', () => {
  it('never blames the API key when the key is working (F11 regression)', () => {
    const msg = describeLlmFailureForWrite(classifyLlmFailure(402, REAL_402, true));
    expect(msg).not.toMatch(/OPENROUTER_API_KEY/);
    expect(msg).toMatch(/out of credits/i);
  });

  it('states plainly that nothing was changed', () => {
    for (const status of [402, 429, 503]) {
      const msg = describeLlmFailureForWrite(classifyLlmFailure(status, 'boom', true));
      expect(msg).toMatch(/Nothing was changed/i);
    }
  });

  it('preserves the provider detail so the user sees the real diagnosis', () => {
    const msg = describeLlmFailureForWrite(classifyLlmFailure(402, REAL_402, true));
    expect(msg).toMatch(/can only afford 2671/);
  });

  it('still directs to the env var when nothing is configured', () => {
    const msg = describeLlmFailureForWrite(classifyLlmFailure(undefined, undefined, false));
    expect(msg).toMatch(/OPENROUTER_API_KEY/);
  });
});

describe('describeLlmFailureForStatus', () => {
  it('names the real cause in the streamed status line', () => {
    expect(describeLlmFailureForStatus(classifyLlmFailure(402, REAL_402, true))).toMatch(
      /out of credits/i,
    );
    expect(describeLlmFailureForStatus(classifyLlmFailure(429, 'rate limit', true))).toMatch(
      /rate-limited/i,
    );
  });

  it('does not claim misconfiguration when a key is present', () => {
    expect(
      describeLlmFailureForStatus(classifyLlmFailure(402, REAL_402, true)),
    ).not.toMatch(/OPENROUTER_API_KEY/);
  });
});
