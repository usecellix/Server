import { Module } from '@nestjs/common';
import { ContextCacheService } from './context-cache.service';

/**
 * TASKS.md #74 — split out so both ExcelAiModule (the cache's original owner
 * and writer) and AuditModule (read-only, for PerformanceMetricsService's
 * cache-hit-rate reporting) can import it without a circular module
 * dependency — ExcelAiModule already imports AuditModule, so AuditModule
 * cannot import ExcelAiModule back.
 */
@Module({
  providers: [ContextCacheService],
  exports: [ContextCacheService],
})
export class ContextCacheModule {}
