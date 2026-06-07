import { Body, Controller, Headers, Post, Res } from '@nestjs/common';
import { FastifyReply } from 'fastify';
import { TRACE_ID_HEADER } from '../common/constants/trace-id.constant';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { ConversationRequestDto } from './dto/conversation-request.dto';
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
}
