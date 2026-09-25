import { Module } from '@nestjs/common';
import { AppConfigModule } from '../../config/app-config.module';
import { ConcurrencyLimitService } from './concurrency-limit.service';

/**
 * A singleton `ConcurrencyLimitService` shared between `ExcelAiModule` and
 * `WebChatModule` (TASKS.md #344) — both spend real LLM cost per request, and
 * the per-user in-flight count must be one shared counter across both
 * surfaces, not two independent ones that could each separately admit up to
 * the cap.
 */
@Module({
  imports: [AppConfigModule],
  providers: [ConcurrencyLimitService],
  exports: [ConcurrencyLimitService],
})
export class ConcurrencyLimitModule {}
