import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { ModelRouter } from '../excel-ai/llm/model-router';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';

@Module({
  imports: [AppConfigModule],
  providers: [ModelRouter, OpenRouterService],
  exports: [OpenRouterService, ModelRouter],
})
export class LlmModule {}
