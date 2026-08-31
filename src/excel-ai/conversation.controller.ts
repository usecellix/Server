import { Body, Controller, Get, Headers, Param, Post, Res, UseGuards } from '@nestjs/common';
import { FastifyReply } from 'fastify';
import { TRACE_ID_HEADER } from '../common/constants/trace-id.constant';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AuthGuard } from '../auth/auth.guard';
import { ConversationRequestDto } from './dto/conversation-request.dto';
import { ToolResultDto } from './dto/tool-result.dto';
import { ConversationService } from './services/conversation.service';

/**
 * Go-live gap (Aug 28, 2026): AuthGuard existed (Mongo-backed, OAuth wired) but was
 * applied nowhere — this endpoint spends OpenRouter credits per call and was open to
 * anyone who could reach the URL. The frontend already sends `credentials: 'include'`
 * on every call (useConversation.ts) and AuthGate/LoginPage already exist — both
 * sides were built and neither was connected. See TASKS.md go-live entry.
 */
@UseGuards(AuthGuard)
@Controller('excel-ai')
export class ConversationController {
  constructor(private readonly conversationService: ConversationService) {}

  @Post('conversation')
  @SkipEnvelope()
  async conversation(
    @Body() body: ConversationRequestDto,
    @Headers(TRACE_ID_HEADER) traceId: string | undefined,
    @Res() reply: FastifyReply,
  ): Promise<void> {
    await this.conversationService.handleConversation(body, reply, traceId);
  }

  @Get('conversation')
  @SkipEnvelope()
  conversationRoot() {
    return {
      ok: true,
      message:
        'Use POST /excel-ai/conversation to send messages, or GET /excel-ai/conversation/:conversationId to load a conversation.',
    };
  }

  @Get('conversation/:conversationId')
  @SkipEnvelope()
  async getConversation(@Param('conversationId') conversationId: string) {
    return this.conversationService.getConversation(conversationId);
  }

  @Post('conversation/tool-result')
  async toolResult(@Body() body: ToolResultDto): Promise<{ accepted: boolean }> {
    return this.conversationService.handleToolResult(body);
  }
}
