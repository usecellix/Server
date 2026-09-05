import { Injectable } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';
import { CHITCHAT_PERSONA_SYSTEM_PROMPT } from '../prompts/chitchat-prompt';
import { LlmCallTelemetry, OpenRouterChatMessage, OpenRouterService } from './openrouter.service';

/**
 * Handles messages LlmRouterService.classifyIntent labeled CHITCHAT — greetings,
 * small talk, identity questions. Deliberately separate from Tier0DirectService:
 * chitchat needs no WorkbookContext at all (no sheet data, no active sheet), so
 * cramming it into a service built around WorkbookContext would mean threading a
 * fake/empty context through for no reason.
 */
@Injectable()
export class ChitchatService {
  constructor(
    private readonly openRouter: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  /** One LOW-tier call, streamed token-by-token for the SSE `chunk` event. */
  streamReply(message: string, telemetry?: LlmCallTelemetry): AsyncGenerator<string> {
    const messages: OpenRouterChatMessage[] = [
      { role: 'system', content: CHITCHAT_PERSONA_SYSTEM_PROMPT },
      { role: 'user', content: message },
    ];

    return this.openRouter.streamChat(messages, telemetry, this.config.openRouterModelLow, 256);
  }
}
