import { Body, Controller, Get, Headers, Param, Post, Res } from '@nestjs/common';
import { FastifyReply } from 'fastify';
import { TRACE_ID_HEADER } from '../common/constants/trace-id.constant';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { ConversationRequestDto } from './dto/conversation-request.dto';
import { ToolResultDto } from './dto/tool-result.dto';
import { ConversationService } from './services/conversation.service';

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
