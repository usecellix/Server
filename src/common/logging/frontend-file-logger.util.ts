import { TELEMETRY_CATEGORIES } from './dto/frontend-log-batch.dto';
/** Retention window for frontend log lines (24 hours) — matches request/planner files. */
export const FRONTEND_LOG_RETENTION_MS = 24 * 60 * 60 * 1000;

export type FrontendLogLevel = 'error' | 'warn' | 'info' | 'action';
/**
 * Derived from the DTO's allowlist rather than restated — TASKS.md #159.
 *
 * This list previously existed FOUR times (client union, request DTO, this
 * file-logger type, and the Mongo schema), independently maintained. Adding
 * `'verify'` in one place 400'd every telemetry batch containing it, and since
 * the DTO rejects a whole batch, `accept.success` was silently lost alongside
 * it. Deriving keeps the compiler honest; `test/telemetry-category-parity.spec.ts`
 * covers the one copy that cannot be derived (the client, across the repo
 * boundary).
 */
export type FrontendLogCategory = (typeof TELEMETRY_CATEGORIES)[number];

export interface FrontendLogEntry {
  ts: string;
  level: FrontendLogLevel;
  category: FrontendLogCategory;
  event: string;
  message: string;
  conversationId?: string;
  changeSetId?: string;
  sessionId?: string;
  workbookKey?: string;
  userAgent?: string;
  pageUrl?: string;
  details?: unknown;
}

export function formatFrontendLogLine(entry: FrontendLogEntry): string {
  return `${JSON.stringify(entry)}\n`;
}

export function pruneFrontendLogLines(
  content: string,
  nowMs: number = Date.now(),
  retentionMs: number = FRONTEND_LOG_RETENTION_MS,
): string {
  const cutoff = nowMs - retentionMs;
  const kept: string[] = [];

  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    try {
      const parsed = JSON.parse(trimmed) as { ts?: unknown };
      if (typeof parsed.ts !== 'string') continue;
      const ms = Date.parse(parsed.ts);
      if (!Number.isFinite(ms) || ms < cutoff) continue;
      kept.push(trimmed);
    } catch {
      // drop bad lines
    }
  }

  return kept.length > 0 ? `${kept.join('\n')}\n` : '';
}
