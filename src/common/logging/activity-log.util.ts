/**
 * Pure helpers for the user-activity log (login, register, AI requests, billing…).
 * No I/O here so the rules are unit-testable.
 */

export const DAY_MS = 24 * 60 * 60 * 1000;
export const DEFAULT_ACTIVITY_RETAIN_DAYS = 3;

export interface ActivityLogEntry {
  ts: string;
  /** Short label, e.g. `login`, `register`, `ai_request`. */
  action: string;
  method: string;
  /** Path only — the query string is never logged (it can carry tokens). */
  path: string;
  statusCode: number;
  responseTimeMs: number;
  ip?: string;
  userAgent?: string;
  userId?: string;
  /** Only filled for the signed-in user, or for login/register attempts. */
  email?: string;
  reqId?: string;
}

const MUTATING = new Set(['POST', 'PUT', 'PATCH', 'DELETE']);

/** Route prefixes (no `/api` prefix) whose mutating requests are user actions. */
const ACTION_PREFIXES: Array<[prefix: string, label: string]> = [
  ['/excel-ai', 'ai_request'],
  ['/web-chat', 'web_chat'],
  ['/billing', 'billing'],
  ['/webhooks', 'webhook'],
  ['/ingest', 'import'],
  ['/bank', 'bank_recon'],
  ['/gst', 'gst_recon'],
  ['/audit', 'audit'],
  ['/sheets-link', 'sheets_link'],
  ['/sheets', 'sheets'],
];

/** Drop the query string and collapse ids so one route is one label. */
export function normalizePath(url: string): string {
  const pathOnly = url.split('?')[0] || '/';
  return pathOnly
    .split('/')
    .map((seg) =>
      /^[0-9a-f]{24}$/i.test(seg) ||
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(seg) ||
      /^\d{4,}$/.test(seg)
        ? ':id'
        : seg,
    )
    .join('/');
}

/**
 * Returns the action label for a request worth logging, or `null` to skip it
 * (health checks, session polling, telemetry, static GETs, CORS preflights).
 */
export function classifyActivity(
  method: string,
  url: string,
  statusCode: number,
): string | null {
  const m = method.toUpperCase();
  const path = normalizePath(url);

  // Better Auth lives under /api/auth/*
  if (path.startsWith('/api/auth/')) {
    const rest = path.slice('/api/auth/'.length);
    if (m === 'POST' && rest.startsWith('sign-in/email')) return 'login';
    if (m === 'POST' && rest.startsWith('sign-in/social')) return 'login_google_start';
    if (rest.startsWith('callback/')) return 'login_oauth_callback';
    if (m === 'POST' && rest.startsWith('sign-up')) return 'register';
    if (m === 'POST' && rest.startsWith('sign-out')) return 'logout';
    if (MUTATING.has(m)) return `auth_${rest.split('/')[0].replace(/[^a-z0-9-]/gi, '') || 'other'}`.replace(/-/g, '_');
    return null; // get-session, error pages, …
  }

  if (!MUTATING.has(m)) return null;

  if (path.startsWith('/excel-login/')) {
    // The task pane polls /claim while the user signs in; 202 = "not ready yet".
    if (statusCode === 202) return null;
    return `excel_login_${path.slice('/excel-login/'.length).split('/')[0] || 'other'}`;
  }

  for (const [prefix, label] of ACTION_PREFIXES) {
    if (path === prefix || path.startsWith(`${prefix}/`)) return label;
  }
  return null;
}

export function clipText(value: unknown, max: number): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.length <= max ? trimmed : trimmed.slice(0, max);
}

/** First hop of X-Forwarded-For (Vercel / proxy), else the socket address. */
export function clientIp(
  forwardedFor: string | string[] | undefined,
  socketIp: string | undefined,
): string | undefined {
  const raw = Array.isArray(forwardedFor) ? forwardedFor[0] : forwardedFor;
  const first = raw?.split(',')[0]?.trim();
  return clipText(first || socketIp, 64);
}

/** `activity-2026-11-13.log` — the UTC date is the rotation key. */
export function activityFileName(at: Date | number = Date.now()): string {
  const d = typeof at === 'number' ? new Date(at) : at;
  return `activity-${d.toISOString().slice(0, 10)}.log`;
}

const FILE_RE = /^activity-(\d{4}-\d{2}-\d{2})\.log$/;

/**
 * A daily file is expired once its whole day is older than `retainDays` —
 * nothing younger than the window is ever deleted.
 */
export function isActivityFileExpired(
  fileName: string,
  nowMs: number = Date.now(),
  retainDays: number = DEFAULT_ACTIVITY_RETAIN_DAYS,
): boolean {
  const match = FILE_RE.exec(fileName);
  if (!match) return false; // never touch files that aren't ours
  const dayStart = Date.parse(`${match[1]}T00:00:00.000Z`);
  if (!Number.isFinite(dayStart)) return false;
  return dayStart + DAY_MS < nowMs - retainDays * DAY_MS;
}

export function formatActivityLine(entry: ActivityLogEntry): string {
  return `${JSON.stringify(entry)}\n`;
}
