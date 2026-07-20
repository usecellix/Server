/** Retention window for request log lines (24 hours). */
export const REQUEST_LOG_RETENTION_MS = 24 * 60 * 60 * 1000;

export interface RequestLogEntry {
  ts: string;
  method: string;
  url: string;
  statusCode: number;
  responseTimeMs: number;
  reqId?: string;
  traceId?: string;
  message?: string;
  /** Captured JSON / SSE / text response (may be truncated). */
  response?: unknown;
}

/**
 * Keep only lines whose leading ISO timestamp is within the retention window.
 * Lines without a parseable timestamp are dropped.
 */
export function pruneRequestLogLines(
  content: string,
  nowMs: number = Date.now(),
  retentionMs: number = REQUEST_LOG_RETENTION_MS,
): string {
  const cutoff = nowMs - retentionMs;
  const lines = content.split(/\r?\n/);
  const kept: string[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    const ts = extractLineTimestamp(trimmed);
    if (ts === null) continue;
    if (ts >= cutoff) {
      kept.push(trimmed);
    }
  }

  return kept.length > 0 ? `${kept.join('\n')}\n` : '';
}

function extractLineTimestamp(line: string): number | null {
  // NDJSON: {"ts":"2026-07-16T13:45:00.000Z",...}
  if (line.startsWith('{')) {
    try {
      const parsed = JSON.parse(line) as { ts?: unknown };
      if (typeof parsed.ts === 'string') {
        const ms = Date.parse(parsed.ts);
        return Number.isFinite(ms) ? ms : null;
      }
    } catch {
      return null;
    }
  }

  // Plain prefix: 2026-07-16T13:45:00.000Z ...
  const match = /^(\d{4}-\d{2}-\d{2}T[\d:.]+Z?)\b/.exec(line);
  if (!match) return null;
  const ms = Date.parse(match[1]);
  return Number.isFinite(ms) ? ms : null;
}

export function formatRequestLogLine(entry: RequestLogEntry): string {
  return `${JSON.stringify(entry)}\n`;
}
