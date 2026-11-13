import { Injectable, Logger, OnModuleDestroy, OnModuleInit } from '@nestjs/common';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';
import {
  activityFileName,
  clientIp,
  clipText,
  classifyActivity,
  DEFAULT_ACTIVITY_RETAIN_DAYS,
  formatActivityLine,
  isActivityFileExpired,
  normalizePath,
  type ActivityLogEntry,
} from './activity-log.util';

const CLEANUP_INTERVAL_MS = 60 * 60 * 1000;

/** The bits of a Fastify request/reply this logger reads. */
export interface ActivityRequestLike {
  method: string;
  url: string;
  id?: unknown;
  ip?: string;
  headers: Record<string, string | string[] | undefined>;
  body?: unknown;
  userSession?: { user?: { id?: string; email?: string | null } };
}

/**
 * User-activity log: one NDJSON file per UTC day (`logs/activity-YYYY-MM-DD.log`).
 * Rotation is the day boundary; files older than ACTIVITY_LOG_RETAIN_DAYS
 * (default 3) are deleted hourly and on startup. Bodies are never written.
 */
@Injectable()
export class ActivityLoggerService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(ActivityLoggerService.name);
  private readonly dir: string;
  private readonly retainDays: number;
  private writeChain: Promise<void> = Promise.resolve();
  private timer: ReturnType<typeof setInterval> | null = null;

  constructor() {
    this.dir =
      process.env.ACTIVITY_LOG_DIR?.trim() ||
      process.env.REQUEST_LOG_DIR?.trim() ||
      path.join(process.cwd(), 'logs');
    const days = Number(process.env.ACTIVITY_LOG_RETAIN_DAYS);
    this.retainDays = Number.isFinite(days) && days >= 1 ? Math.floor(days) : DEFAULT_ACTIVITY_RETAIN_DAYS;
  }

  async onModuleInit(): Promise<void> {
    await fs.mkdir(this.dir, { recursive: true }).catch(() => undefined);
    await this.cleanup();
    this.timer = setInterval(() => void this.cleanup(), CLEANUP_INTERVAL_MS);
    if (typeof this.timer.unref === 'function') this.timer.unref();
    this.logger.log(`Activity log → ${this.dir}/activity-YYYY-MM-DD.log (retain ${this.retainDays}d)`);
  }

  onModuleDestroy(): void {
    if (this.timer) clearInterval(this.timer);
    this.timer = null;
  }

  /** Call once per completed response; non-user-action requests are ignored. */
  record(request: ActivityRequestLike, statusCode: number, responseTimeMs: number): void {
    const action = classifyActivity(request.method, request.url, statusCode);
    if (!action) return;

    const user = request.userSession?.user;
    // Login/register have no session yet — take the attempted email from the
    // body. The password and every other field are never read.
    const attemptedEmail =
      action === 'login' || action === 'register'
        ? clipText((request.body as { email?: unknown } | undefined)?.email, 254)
        : undefined;
    const uaHeader = request.headers['user-agent'];
    const ip = clientIp(request.headers['x-forwarded-for'], request.ip);
    const userAgent = clipText(Array.isArray(uaHeader) ? uaHeader[0] : uaHeader, 160);
    const email = user?.email ?? attemptedEmail;

    const entry: ActivityLogEntry = {
      ts: new Date().toISOString(),
      action,
      method: request.method,
      path: normalizePath(request.url),
      statusCode,
      responseTimeMs,
      ...(ip ? { ip } : {}),
      ...(userAgent ? { userAgent } : {}),
      ...(user?.id ? { userId: user.id } : {}),
      ...(email ? { email } : {}),
      ...(request.id !== undefined && request.id !== '' ? { reqId: String(request.id) } : {}),
    };

    this.writeChain = this.writeChain
      .then(async () => {
        await fs.mkdir(this.dir, { recursive: true });
        await fs.appendFile(path.join(this.dir, activityFileName()), formatActivityLine(entry), 'utf8');
      })
      .catch((err: unknown) => {
        this.logger.warn(`Failed to write activity log: ${err instanceof Error ? err.message : String(err)}`);
      });
  }

  /** Delete daily files that are entirely older than the retention window. */
  async cleanup(now: number = Date.now()): Promise<string[]> {
    const removed: string[] = [];
    try {
      for (const name of await fs.readdir(this.dir)) {
        if (!isActivityFileExpired(name, now, this.retainDays)) continue;
        await fs.unlink(path.join(this.dir, name));
        removed.push(name);
      }
    } catch (err: unknown) {
      if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') {
        this.logger.warn(`Activity log cleanup failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
    return removed;
  }

  /** Exposed for tests. */
  async flush(): Promise<void> {
    await this.writeChain;
  }
}
