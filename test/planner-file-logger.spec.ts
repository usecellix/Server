import {
  formatPlannerLogLine,
  PLANNER_LOG_TEXT_MAX,
  truncateForPlannerLog,
} from '../src/common/logging/planner-file-logger.util';
import { pruneRequestLogLines, REQUEST_LOG_RETENTION_MS } from '../src/common/logging/request-file-logger.util';

describe('planner-file-logger.util', () => {
  it('formats NDJSON planner entries', () => {
    const line = formatPlannerLogLine({
      ts: '2026-07-18T10:00:00.000Z',
      correlationId: 'req_1',
      model: 'test-model',
      durationMs: 100,
      success: true,
      input: {
        prompt: 'sort by amount',
        userMessage: 'user…',
        historyLength: 0,
        sheets: ['Sheet1'],
        activeSheet: 'Sheet1',
        hasPromptContext: false,
      },
      output: {
        raw: '{"subtasks":[]}',
        parsed: { subtasks: [], clarificationsNeeded: [], confidence: 'high', reasoning: 'ok' },
        fallback: false,
        retried: false,
      },
    });

    expect(line.endsWith('\n')).toBe(true);
    const parsed = JSON.parse(line.trim()) as { input: { prompt: string }; output: { fallback: boolean } };
    expect(parsed.input.prompt).toBe('sort by amount');
    expect(parsed.output.fallback).toBe(false);
  });

  it('truncates long text for the log file', () => {
    const long = 'x'.repeat(PLANNER_LOG_TEXT_MAX + 5_000);
    const result = truncateForPlannerLog(long);
    expect(result.truncated).toBe(true);
    expect(result.value.length).toBeLessThan(long.length);
    expect(result.value).toContain('[truncated');
  });

  it('prunes old planner lines via shared retention helper', () => {
    const now = Date.parse('2026-07-18T12:00:00.000Z');
    const oldTs = new Date(now - REQUEST_LOG_RETENTION_MS - 60_000).toISOString();
    const freshTs = new Date(now - 1_000).toISOString();

    const content = [
      formatPlannerLogLine({
        ts: oldTs,
        correlationId: 'old',
        model: 'm',
        durationMs: 1,
        success: true,
        input: {
          prompt: 'old',
          userMessage: 'old',
          historyLength: 0,
          sheets: [],
          activeSheet: 'Sheet1',
          hasPromptContext: false,
        },
        output: { raw: '{}', parsed: {}, fallback: false, retried: false },
      }).trim(),
      formatPlannerLogLine({
        ts: freshTs,
        correlationId: 'fresh',
        model: 'm',
        durationMs: 1,
        success: true,
        input: {
          prompt: 'fresh',
          userMessage: 'fresh',
          historyLength: 0,
          sheets: [],
          activeSheet: 'Sheet1',
          hasPromptContext: false,
        },
        output: { raw: '{}', parsed: {}, fallback: false, retried: false },
      }).trim(),
    ].join('\n');

    const pruned = pruneRequestLogLines(content, now);
    expect(pruned).not.toContain('"correlationId":"old"');
    expect(pruned).toContain('"correlationId":"fresh"');
  });
});
