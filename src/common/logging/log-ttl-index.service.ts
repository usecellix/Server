import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { InjectConnection } from '@nestjs/mongoose';
import { Connection } from 'mongoose';
import { LOG_TTL_SECONDS } from './schemas/request-log.schema';

/**
 * Ensures Mongo TTL indexes exist on request_logs / planner_logs.
 * Drops a conflicting non-TTL `ts_1` index if present so expireAfterSeconds can be applied.
 */
@Injectable()
export class LogTtlIndexService implements OnModuleInit {
  private readonly logger = new Logger(LogTtlIndexService.name);

  constructor(@InjectConnection() private readonly connection: Connection) {}

  async onModuleInit(): Promise<void> {
    await this.ensureTtl('request_logs');
    await this.ensureTtl('planner_logs');
    await this.ensureTtl('frontend_logs');
  }

  private async ensureTtl(collectionName: string): Promise<void> {
    try {
      const col = this.connection.collection(collectionName);
      const indexes = await col.indexes();
      const tsIndex = indexes.find(
        (idx) =>
          idx.key &&
          Object.keys(idx.key).length === 1 &&
          (idx.key as Record<string, number>).ts === 1,
      );

      if (tsIndex && tsIndex.expireAfterSeconds !== LOG_TTL_SECONDS) {
        if (tsIndex.name) {
          this.logger.warn(
            `Replacing index ${tsIndex.name} on ${collectionName} with 3-day TTL`,
          );
          await col.dropIndex(tsIndex.name);
        }
      }

      await col.createIndex({ ts: 1 }, { expireAfterSeconds: LOG_TTL_SECONDS });
      this.logger.log(
        `${collectionName}: TTL index on ts (${LOG_TTL_SECONDS}s / 3 days)`,
      );
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to ensure TTL on ${collectionName}: ${msg}`);
    }
  }
}
