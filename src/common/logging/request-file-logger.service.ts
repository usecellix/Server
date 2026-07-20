import { Injectable, Logger, OnModuleDestroy, OnModuleInit } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';
import {
  formatRequestLogLine,
  pruneRequestLogLines,
  REQUEST_LOG_RETENTION_MS,
  type RequestLogEntry,
} from './request-file-logger.util';
import { RequestLog } from './schemas/request-log.schema';

const DEFAULT_LOG_DIR = path.join(process.cwd(), 'logs');
const DEFAULT_LOG_FILE = 'requests.log';
/** Prune at most once per hour (also runs on startup). */
const PRUNE_INTERVAL_MS = 60 * 60 * 1000;

export type RequestFileLogInput = Omit<RequestLogEntry, 'ts'> & { ts?: string };

@Injectable()
export class RequestFileLoggerService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(RequestFileLoggerService.name);
  private readonly filePath: string;
  private writeChain: Promise<void> = Promise.resolve();
  private pruneTimer: ReturnType<typeof setInterval> | null = null;
  private lastPruneAt = 0;

  constructor(
    @InjectModel(RequestLog.name)
    private readonly requestLogModel: Model<RequestLog>,
  ) {
    const dir = process.env.REQUEST_LOG_DIR?.trim() || DEFAULT_LOG_DIR;
    const file = process.env.REQUEST_LOG_FILE?.trim() || DEFAULT_LOG_FILE;
    this.filePath = path.isAbsolute(file) ? file : path.join(dir, file);
  }

  async onModuleInit(): Promise<void> {
    await this.ensureLogDir();
    await this.pruneIfDue(true);
    this.pruneTimer = setInterval(() => {
      void this.pruneIfDue(true);
    }, PRUNE_INTERVAL_MS);
    // Allow process to exit without waiting for the interval.
    if (typeof this.pruneTimer.unref === 'function') {
      this.pruneTimer.unref();
    }
    this.logger.log(`Request file logger → ${this.filePath} (retain ${REQUEST_LOG_RETENTION_MS / 3_600_000}h)`);
  }

  onModuleDestroy(): void {
    if (this.pruneTimer) {
      clearInterval(this.pruneTimer);
      this.pruneTimer = null;
    }
  }

  /** Append one completed HTTP request; old lines (>24h) are dropped from the top periodically. */
  logRequest(input: RequestFileLogInput): void {
    const entry: RequestLogEntry = {
      ts: input.ts ?? new Date().toISOString(),
      method: input.method,
      url: input.url,
      statusCode: input.statusCode,
      responseTimeMs: input.responseTimeMs,
      ...(input.reqId ? { reqId: input.reqId } : {}),
      ...(input.traceId ? { traceId: input.traceId } : {}),
      ...(input.message ? { message: input.message } : {}),
      ...(input.response !== undefined ? { response: input.response } : {}),
    };

    this.writeChain = this.writeChain
      .then(async () => {
        await this.ensureLogDir();
        await fs.appendFile(this.filePath, formatRequestLogLine(entry), 'utf8');
        await this.pruneIfDue(false);
        await this.persistToDb(entry);
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        this.logger.warn(`Failed to write request log: ${msg}`);
      });
  }

  getLogFilePath(): string {
    return this.filePath;
  }

  /** Exposed for tests. */
  async flush(): Promise<void> {
    await this.writeChain;
  }

  private async persistToDb(entry: RequestLogEntry): Promise<void> {
    try {
      await this.requestLogModel.create({
        ts: new Date(entry.ts),
        method: entry.method,
        url: entry.url,
        statusCode: entry.statusCode,
        responseTimeMs: entry.responseTimeMs,
        ...(entry.reqId ? { reqId: entry.reqId } : {}),
        ...(entry.traceId ? { traceId: entry.traceId } : {}),
        ...(entry.message ? { message: entry.message } : {}),
        ...(entry.response !== undefined ? { response: entry.response } : {}),
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to persist request log to MongoDB: ${msg}`);
    }
  }

  private async ensureLogDir(): Promise<void> {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
  }

  private async pruneIfDue(force: boolean): Promise<void> {
    const now = Date.now();
    if (!force && now - this.lastPruneAt < PRUNE_INTERVAL_MS) {
      return;
    }
    this.lastPruneAt = now;

    try {
      let raw: string;
      try {
        raw = await fs.readFile(this.filePath, 'utf8');
      } catch (err: unknown) {
        if ((err as NodeJS.ErrnoException)?.code === 'ENOENT') {
          return;
        }
        throw err;
      }

      const pruned = pruneRequestLogLines(raw, now, REQUEST_LOG_RETENTION_MS);
      if (pruned === raw) {
        return;
      }

      const tmpPath = `${this.filePath}.tmp`;
      await fs.writeFile(tmpPath, pruned, 'utf8');
      await fs.rename(tmpPath, this.filePath);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to prune request log: ${msg}`);
    }
  }
}
