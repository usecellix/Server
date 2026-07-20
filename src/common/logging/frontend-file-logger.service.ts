import { Injectable, Logger, OnModuleDestroy, OnModuleInit } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';
import {
  formatFrontendLogLine,
  FRONTEND_LOG_RETENTION_MS,
  pruneFrontendLogLines,
  type FrontendLogEntry,
} from './frontend-file-logger.util';
import { FrontendLog } from './schemas/frontend-log.schema';

const DEFAULT_LOG_DIR = path.join(process.cwd(), 'logs');
const DEFAULT_LOG_FILE = 'frontend.log';
const PRUNE_INTERVAL_MS = 60 * 60 * 1000;

export type FrontendFileLogInput = Omit<FrontendLogEntry, 'ts'> & { ts?: string };

@Injectable()
export class FrontendFileLoggerService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(FrontendFileLoggerService.name);
  private readonly filePath: string;
  private writeChain: Promise<void> = Promise.resolve();
  private pruneTimer: ReturnType<typeof setInterval> | null = null;
  private lastPruneAt = 0;

  constructor(
    @InjectModel(FrontendLog.name)
    private readonly frontendLogModel: Model<FrontendLog>,
  ) {
    const dir = process.env.FRONTEND_LOG_DIR?.trim() || DEFAULT_LOG_DIR;
    const file = process.env.FRONTEND_LOG_FILE?.trim() || DEFAULT_LOG_FILE;
    this.filePath = path.isAbsolute(file) ? file : path.join(dir, file);
  }

  async onModuleInit(): Promise<void> {
    await this.ensureLogDir();
    await this.pruneIfDue(true);
    this.pruneTimer = setInterval(() => {
      void this.pruneIfDue(true);
    }, PRUNE_INTERVAL_MS);
    if (typeof this.pruneTimer.unref === 'function') {
      this.pruneTimer.unref();
    }
    this.logger.log(
      `Frontend file logger → ${this.filePath} (retain ${FRONTEND_LOG_RETENTION_MS / 3_600_000}h)`,
    );
  }

  onModuleDestroy(): void {
    if (this.pruneTimer) {
      clearInterval(this.pruneTimer);
      this.pruneTimer = null;
    }
  }

  logEvent(input: FrontendFileLogInput): void {
    const entry: FrontendLogEntry = {
      ts: input.ts ?? new Date().toISOString(),
      level: input.level,
      category: input.category,
      event: input.event,
      message: input.message,
      ...(input.conversationId ? { conversationId: input.conversationId } : {}),
      ...(input.changeSetId ? { changeSetId: input.changeSetId } : {}),
      ...(input.sessionId ? { sessionId: input.sessionId } : {}),
      ...(input.workbookKey ? { workbookKey: input.workbookKey } : {}),
      ...(input.userAgent ? { userAgent: input.userAgent } : {}),
      ...(input.pageUrl ? { pageUrl: input.pageUrl } : {}),
      ...(input.details !== undefined ? { details: input.details } : {}),
    };

    this.writeChain = this.writeChain
      .then(async () => {
        await this.ensureLogDir();
        await fs.appendFile(this.filePath, formatFrontendLogLine(entry), 'utf8');
        await this.pruneIfDue(false);
        await this.persistToDb(entry);
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        this.logger.warn(`Failed to write frontend log: ${msg}`);
      });
  }

  logEvents(inputs: FrontendFileLogInput[]): void {
    for (const input of inputs) {
      this.logEvent(input);
    }
  }

  async flush(): Promise<void> {
    await this.writeChain;
  }

  getLogFilePath(): string {
    return this.filePath;
  }

  private async persistToDb(entry: FrontendLogEntry): Promise<void> {
    try {
      await this.frontendLogModel.create({
        ts: new Date(entry.ts),
        level: entry.level,
        category: entry.category,
        event: entry.event,
        message: entry.message,
        ...(entry.conversationId ? { conversationId: entry.conversationId } : {}),
        ...(entry.changeSetId ? { changeSetId: entry.changeSetId } : {}),
        ...(entry.sessionId ? { sessionId: entry.sessionId } : {}),
        ...(entry.workbookKey ? { workbookKey: entry.workbookKey } : {}),
        ...(entry.userAgent ? { userAgent: entry.userAgent } : {}),
        ...(entry.pageUrl ? { pageUrl: entry.pageUrl } : {}),
        ...(entry.details !== undefined ? { details: entry.details } : {}),
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to persist frontend log to MongoDB: ${msg}`);
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

      const pruned = pruneFrontendLogLines(raw, now, FRONTEND_LOG_RETENTION_MS);
      if (pruned === raw) {
        return;
      }

      const tmpPath = `${this.filePath}.tmp`;
      await fs.writeFile(tmpPath, pruned, 'utf8');
      await fs.rename(tmpPath, this.filePath);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to prune frontend log: ${msg}`);
    }
  }
}
