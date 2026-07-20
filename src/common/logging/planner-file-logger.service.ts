import { Injectable, Logger, OnModuleDestroy, OnModuleInit } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { promises as fs } from 'node:fs';
import * as path from 'node:path';
import {
  pruneRequestLogLines,
  REQUEST_LOG_RETENTION_MS,
} from './request-file-logger.util';
import {
  formatPlannerLogLine,
  type PlannerLogEntry,
} from './planner-file-logger.util';
import { PlannerLog } from './schemas/planner-log.schema';

const DEFAULT_LOG_DIR = path.join(process.cwd(), 'logs');
const DEFAULT_LOG_FILE = 'planner.log';
/** Prune at most once per hour (also runs on startup). */
const PRUNE_INTERVAL_MS = 60 * 60 * 1000;

export type PlannerFileLogInput = Omit<PlannerLogEntry, 'ts'> & { ts?: string };

@Injectable()
export class PlannerFileLoggerService implements OnModuleInit, OnModuleDestroy {
  private readonly logger = new Logger(PlannerFileLoggerService.name);
  private readonly filePath: string;
  private writeChain: Promise<void> = Promise.resolve();
  private pruneTimer: ReturnType<typeof setInterval> | null = null;
  private lastPruneAt = 0;

  constructor(
    @InjectModel(PlannerLog.name)
    private readonly plannerLogModel: Model<PlannerLog>,
  ) {
    const dir = process.env.PLANNER_LOG_DIR?.trim() || process.env.REQUEST_LOG_DIR?.trim() || DEFAULT_LOG_DIR;
    const file = process.env.PLANNER_LOG_FILE?.trim() || DEFAULT_LOG_FILE;
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
      `Planner file logger → ${this.filePath} (retain ${REQUEST_LOG_RETENTION_MS / 3_600_000}h)`,
    );
  }

  onModuleDestroy(): void {
    if (this.pruneTimer) {
      clearInterval(this.pruneTimer);
      this.pruneTimer = null;
    }
  }

  /** Append one planner call (input + output) as NDJSON. */
  logPlanner(input: PlannerFileLogInput): void {
    const entry: PlannerLogEntry = {
      ts: input.ts ?? new Date().toISOString(),
      correlationId: input.correlationId,
      model: input.model,
      durationMs: input.durationMs,
      success: input.success,
      ...(input.error ? { error: input.error } : {}),
      input: input.input,
      output: input.output,
    };

    this.writeChain = this.writeChain
      .then(async () => {
        await this.ensureLogDir();
        await fs.appendFile(this.filePath, formatPlannerLogLine(entry), 'utf8');
        await this.pruneIfDue(false);
        await this.persistToDb(entry);
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        this.logger.warn(`Failed to write planner log: ${msg}`);
      });
  }

  getLogFilePath(): string {
    return this.filePath;
  }

  /** True when PLANNER_LOG_FULL_PROMPTS=true — include system prompt in log lines. */
  shouldLogFullPrompts(): boolean {
    return process.env.PLANNER_LOG_FULL_PROMPTS?.trim().toLowerCase() === 'true';
  }

  /** Exposed for tests. */
  async flush(): Promise<void> {
    await this.writeChain;
  }

  private async persistToDb(entry: PlannerLogEntry): Promise<void> {
    try {
      await this.plannerLogModel.create({
        ts: new Date(entry.ts),
        correlationId: entry.correlationId,
        model: entry.model,
        durationMs: entry.durationMs,
        success: entry.success,
        ...(entry.error ? { error: entry.error } : {}),
        input: entry.input as unknown as Record<string, unknown>,
        output: entry.output as unknown as Record<string, unknown>,
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to persist planner log to MongoDB: ${msg}`);
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
      this.logger.warn(`Failed to prune planner log: ${msg}`);
    }
  }
}
