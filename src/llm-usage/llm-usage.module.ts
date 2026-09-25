import { Global, Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { LlmUsageService } from './llm-usage.service';
import { AiPrompt, AiPromptSchema } from './schemas/ai-prompt.schema';
import { LlmCall, LlmCallSchema } from './schemas/llm-call.schema';

/** Global so every module that provides its own OpenRouterService can record usage. */
@Global()
@Module({
  imports: [
    MongooseModule.forFeature([
      { name: AiPrompt.name, schema: AiPromptSchema },
      { name: LlmCall.name, schema: LlmCallSchema },
    ]),
  ],
  providers: [LlmUsageService],
  exports: [LlmUsageService],
})
export class LlmUsageModule {}
