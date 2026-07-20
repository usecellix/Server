import {
  formatRequestLogLine,
  pruneRequestLogLines,
  REQUEST_LOG_RETENTION_MS,
} from '../src/common/logging/request-file-logger.util';
import {
  summarizeResponseForLog,
  type CapturedResponse,
} from '../src/common/logging/request-response-capture.util';

describe('request-file-logger.util', () => {
  const now = Date.parse('2026-07-16T12:00:00.000Z');

  it('keeps lines within the last 24 hours and drops older ones from the top', () => {
    const oldTs = new Date(now - REQUEST_LOG_RETENTION_MS - 60_000).toISOString();
    const midTs = new Date(now - 3_600_000).toISOString();
    const freshTs = new Date(now - 1_000).toISOString();

    const content = [
      formatRequestLogLine({
        ts: oldTs,
        method: 'POST',
        url: '/old',
        statusCode: 200,
        responseTimeMs: 10,
      }).trim(),
      formatRequestLogLine({
        ts: midTs,
        method: 'GET',
        url: '/mid',
        statusCode: 200,
        responseTimeMs: 5,
      }).trim(),
      formatRequestLogLine({
        ts: freshTs,
        method: 'POST',
        url: '/fresh',
        statusCode: 200,
        responseTimeMs: 20,
      }).trim(),
      '',
    ].join('\n');

    const pruned = pruneRequestLogLines(content, now);
    expect(pruned).not.toContain('/old');
    expect(pruned).toContain('/mid');
    expect(pruned).toContain('/fresh');
    expect(pruned.startsWith('{')).toBe(true);
  });

  it('returns empty string when all lines are older than retention', () => {
    const oldTs = new Date(now - REQUEST_LOG_RETENTION_MS - 1).toISOString();
    const content = formatRequestLogLine({
      ts: oldTs,
      method: 'POST',
      url: '/excel-ai/conversation',
      statusCode: 200,
      responseTimeMs: 100,
    });

    expect(pruneRequestLogLines(content, now)).toBe('');
  });

  it('drops unparseable lines', () => {
    const freshTs = new Date(now).toISOString();
    const content = `not-json\n${formatRequestLogLine({
      ts: freshTs,
      method: 'POST',
      url: '/ok',
      statusCode: 200,
      responseTimeMs: 1,
    })}`;

    const pruned = pruneRequestLogLines(content, now);
    expect(pruned).toContain('/ok');
    expect(pruned).not.toContain('not-json');
  });

  it('includes response field when formatting log lines', () => {
    const line = formatRequestLogLine({
      ts: new Date(now).toISOString(),
      method: 'POST',
      url: '/excel-ai/conversation',
      statusCode: 200,
      responseTimeMs: 50,
      message: 'hello',
      response: { kind: 'sse', events: [{ event: 'answer', data: { answer: 'hi' } }] },
    });
    const parsed = JSON.parse(line) as { response?: { kind?: string } };
    expect(parsed.response?.kind).toBe('sse');
  });
});

describe('summarizeResponseForLog', () => {
  it('returns sse payload as-is when small', () => {
    const response: CapturedResponse = {
      kind: 'sse',
      events: [{ event: 'answer', data: { answer: 'Total is 10' } }],
    };
    expect(summarizeResponseForLog(response)).toEqual(response);
  });

  it('truncates oversized payloads', () => {
    const huge = 'x'.repeat(20_000);
    const response: CapturedResponse = {
      kind: 'json',
      body: { answer: huge },
    };
    const summarized = summarizeResponseForLog(response) as {
      truncated?: boolean;
      kind?: string;
    };
    expect(summarized.truncated).toBe(true);
    expect(summarized.kind).toBe('json');
  });
});
