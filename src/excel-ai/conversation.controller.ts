import {
  Body,
  Controller,
  Delete,
  Get,
  Headers,
  HttpCode,
  Param,
  Patch,
  Post,
  Query,
  Res,
  UseGuards,
} from '@nestjs/common';
import { FastifyReply } from 'fastify';
import { TRACE_ID_HEADER } from '../common/constants/trace-id.constant';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AuthGuard, AuthUserSession, Session } from '../auth/auth.guard';
import { ConversationRequestDto } from './dto/conversation-request.dto';
import { ListConversationsQueryDto } from './dto/list-conversations-query.dto';
import { RenameConversationDto } from './dto/rename-conversation.dto';
import { ToolResultDto } from './dto/tool-result.dto';
import { ContinueRunDto } from './dto/continue-run.dto';
import { ConversationService } from './services/conversation.service';

/**
 * Go-live gap (Aug 28, 2026): AuthGuard existed (Mongo-backed, OAuth wired) but was
 * applied nowhere — this endpoint spends OpenRouter credits per call and was open to
 * anyone who could reach the URL. The frontend already sends `credentials: 'include'`
 * on every call (useConversation.ts) and AuthGate/LoginPage already exist — both
 * sides were built and neither was connected. See TASKS.md go-live entry.
 *
 * The guard runs on every route in this controller, so `@Session()` is always
 * populated — that is what makes it safe for TASKS.md #170's `userId` to come
 * from the session rather than from the request body.
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
    @Session() session: AuthUserSession | undefined,
  ): Promise<void> {
    await this.conversationService.handleConversation(
      body,
      reply,
      traceId,
      session?.user?.id,
    );
  }

  @Get('conversation')
  @SkipEnvelope()
  conversationRoot() {
    return {
      ok: true,
      message:
        'Use POST /excel-ai/conversation to send messages, GET /excel-ai/conversations to list your past conversations, or GET /excel-ai/conversation/:conversationId to load one.',
    };
  }

  /**
   * The signed-in user's past conversations, newest first (TASKS.md #171).
   *
   * Route sits above `conversation/:conversationId` in the file for readability
   * only — it is a distinct path (`conversations`, plural) so there is no
   * ordering hazard between them.
   */
  @Get('conversations')
  @SkipEnvelope()
  async listConversations(
    @Query() query: ListConversationsQueryDto,
    @Session() session: AuthUserSession | undefined,
  ) {
    return this.conversationService.listConversations(session!.user.id, {
      limit: query.limit,
      cursor: query.cursor,
      workbookId: query.workbookId,
    });
  }

  @Get('conversation/:conversationId')
  @SkipEnvelope()
  async getConversation(
    @Param('conversationId') conversationId: string,
    @Session() session: AuthUserSession | undefined,
  ) {
    return this.conversationService.getConversation(conversationId, session?.user?.id);
  }

  /** Rename a conversation (TASKS.md #177) — the history list's CRUD "R". */
  @Patch('conversation/:conversationId')
  @SkipEnvelope()
  async renameConversation(
    @Param('conversationId') conversationId: string,
    @Body() body: RenameConversationDto,
    @Session() session: AuthUserSession | undefined,
  ) {
    return this.conversationService.renameConversation(
      conversationId,
      session!.user.id,
      body.title,
    );
  }

  /** Delete a conversation (TASKS.md #177) — hard delete, not soft-archive. */
  @Delete('conversation/:conversationId')
  @HttpCode(204)
  async deleteConversation(
    @Param('conversationId') conversationId: string,
    @Session() session: AuthUserSession | undefined,
  ): Promise<void> {
    await this.conversationService.deleteConversation(conversationId, session!.user.id);
  }

  @Post('conversation/tool-result')
  async toolResult(@Body() body: ToolResultDto): Promise<{ accepted: boolean }> {
    return this.conversationService.handleToolResult(body);
  }

  /**
   * Advances a step-wise Tier 3 run (TASKS.md #153, STEPWISE_EXECUTION.md §3).
   *
   * Streams exactly like `POST conversation` — the client's decision on the
   * wave it was last shown goes up, the next wave's Accept card comes back, and
   * the stream ends again. `userId` comes from the session so a caller holding
   * someone else's runId cannot drive their build.
   */
  @Post('conversation/continue')
  @SkipEnvelope()
  async continueRun(
    @Body() body: ContinueRunDto,
    @Headers(TRACE_ID_HEADER) traceId: string | undefined,
    @Res() reply: FastifyReply,
    @Session() session: AuthUserSession | undefined,
  ): Promise<void> {
    await this.conversationService.continueRun(body, reply, traceId, session?.user?.id);
  }
}
