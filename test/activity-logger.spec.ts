import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import * as path from 'node:path';
import {
  activityFileName,
  classifyActivity,
  isActivityFileExpired,
  normalizePath,
} from '../src/common/logging/activity-log.util';
import { ActivityLoggerService } from '../src/common/logging/activity-logger.service';

describe('classifyActivity', () => {
  it('labels login, register, logout and oauth', () => {
    expect(classifyActivity('POST', '/api/auth/sign-in/email', 200)).toBe('login');
    expect(classifyActivity('POST', '/api/auth/sign-up/email', 200)).toBe('register');
    expect(classifyActivity('POST', '/api/auth/sign-out', 200)).toBe('logout');
    expect(classifyActivity('POST', '/api/auth/sign-in/social', 200)).toBe('login_google_start');
    expect(classifyActivity('GET', '/api/auth/callback/google?code=x&state=y', 302)).toBe('login_oauth_callback');
  });

  it('labels user requests', () => {
    expect(classifyActivity('POST', '/excel-ai/conversation', 200)).toBe('ai_request');
    expect(classifyActivity('POST', '/billing/checkout/subscribe', 200)).toBe('billing');
    expect(classifyActivity('POST', '/webhooks/razorpay', 200)).toBe('webhook');
  });

  it('skips noise: session polling, health, telemetry, preflight, GETs', () => {
    expect(classifyActivity('GET', '/api/auth/get-session', 200)).toBeNull();
    expect(classifyActivity('GET', '/health', 200)).toBeNull();
    expect(classifyActivity('POST', '/telemetry/frontend', 200)).toBeNull();
    expect(classifyActivity('OPTIONS', '/excel-ai/conversation', 204)).toBeNull();
    expect(classifyActivity('GET', '/excel-ai/anything', 200)).toBeNull();
  });

  it('skips the 202 "not ready" claim polls but logs the real claim', () => {
    expect(classifyActivity('POST', '/excel-login/claim', 202)).toBeNull();
    expect(classifyActivity('POST', '/excel-login/claim', 200)).toBe('excel_login_claim');
  });
});

describe('normalizePath', () => {
  it('drops the query string (tokens) and collapses ids', () => {
    expect(normalizePath('/excel-login/x?token=secret')).toBe('/excel-login/x');
    expect(normalizePath('/audit/507f1f77bcf86cd799439011/revert')).toBe('/audit/:id/revert');
  });
});

describe('retention', () => {
  const now = Date.parse('2026-11-13T12:00:00.000Z');

  it('names files by UTC day', () => {
    expect(activityFileName(now)).toBe('activity-2026-11-13.log');
  });

  it('keeps the last 3 days, expires older, ignores foreign files', () => {
    expect(isActivityFileExpired('activity-2026-11-13.log', now, 3)).toBe(false);
    expect(isActivityFileExpired('activity-2026-11-10.log', now, 3)).toBe(false); // ends 11-11 00:00, within 72h
    expect(isActivityFileExpired('activity-2026-11-09.log', now, 3)).toBe(true);
    expect(isActivityFileExpired('requests.log', now, 3)).toBe(false);
    expect(isActivityFileExpired('activity-notadate.log', now, 3)).toBe(false);
  });
});

describe('ActivityLoggerService', () => {
  let dir: string;
  const saved = { ...process.env };

  beforeEach(async () => {
    dir = await mkdtemp(path.join(tmpdir(), 'activity-'));
    process.env.ACTIVITY_LOG_DIR = dir;
    delete process.env.ACTIVITY_LOG_RETAIN_DAYS;
  });
  afterEach(async () => {
    process.env = { ...saved };
    await rm(dir, { recursive: true, force: true });
  });

  it('writes a login line with the email but never the password or query', async () => {
    const svc = new ActivityLoggerService();
    svc.record(
      {
        method: 'POST',
        url: '/api/auth/sign-in/email?x=1',
        headers: { 'user-agent': 'jest', 'x-forwarded-for': '1.2.3.4, 10.0.0.1' },
        body: { email: 'a@b.com', password: 'hunter2' },
        id: 'req-1',
      },
      200,
      12,
    );
    await svc.flush();
    const raw = await readFile(path.join(dir, activityFileName()), 'utf8');
    const line = JSON.parse(raw.trim());
    expect(line).toMatchObject({ action: 'login', email: 'a@b.com', ip: '1.2.3.4', statusCode: 200, path: '/api/auth/sign-in/email' });
    expect(raw).not.toContain('hunter2');
    expect(raw).not.toContain('x=1');
  });

  it('writes nothing for non-user-action requests', async () => {
    const svc = new ActivityLoggerService();
    svc.record({ method: 'GET', url: '/health', headers: {} }, 200, 1);
    await svc.flush();
    expect(await readdir(dir)).toEqual([]);
  });

  it('cleanup removes expired daily files and leaves everything else', async () => {
    const svc = new ActivityLoggerService();
    await writeFile(path.join(dir, 'activity-2026-11-01.log'), 'old\n');
    await writeFile(path.join(dir, 'activity-2026-11-12.log'), 'recent\n');
    await writeFile(path.join(dir, 'requests.log'), 'other\n');
    const removed = await svc.cleanup(Date.parse('2026-11-13T12:00:00.000Z'));
    expect(removed).toEqual(['activity-2026-11-01.log']);
    expect((await readdir(dir)).sort()).toEqual(['activity-2026-11-12.log', 'requests.log']);
  });
});
