import { StructuredLogger } from '../src/agents/logging/structured-logger';

describe('StructuredLogger', () => {
  it('redacts API keys, bearer tokens and password-like values', () => {
    const logger = new StructuredLogger() as unknown as {
      redactText: (value: string) => string;
    };
    const redacted = logger.redactText(
      'api_key=sk-live-1234567890 token: abc123 Bearer secret-token password=myPass',
    );

    expect(redacted).not.toContain('sk-live-1234567890');
    expect(redacted).not.toContain('abc123');
    expect(redacted).not.toContain('secret-token');
    expect(redacted).not.toContain('myPass');
    expect(redacted).toContain('[REDACTED');
  });

  it('estimates token usage from response length', () => {
    const logger = new StructuredLogger();
    expect(logger.estimateTokens('')).toBe(0);
    expect(logger.estimateTokens('1234')).toBe(1);
    expect(logger.estimateTokens('12345')).toBe(2);
  });

  it('emits tier_decision events with redacted message', () => {
    const logger = new StructuredLogger();
    const logSpy = jest.spyOn(
      (logger as unknown as { logger: { log: (msg: string) => void } }).logger,
      'log',
    );

    logger.logTierDecision({
      traceId: 'trace-1',
      message: 'calculate GST api_key=sk-live-1234567890',
      tier: 2,
      matchedBy: 'regex',
      actionHint: 'FORMULA_GEN',
      llmCallCount: 2,
      durationMs: 850,
    });

    expect(logSpy).toHaveBeenCalledTimes(1);
    const payload = JSON.parse(logSpy.mock.calls[0][0] as string) as Record<string, unknown>;
    expect(payload.event).toBe('tier_decision');
    expect(payload.tier).toBe(2);
    expect(payload.llmCallCount).toBe(2);
    expect(String(payload.message)).not.toContain('sk-live-1234567890');
  });
});
