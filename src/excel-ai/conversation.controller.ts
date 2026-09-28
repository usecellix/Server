import {
  Body,
  Controller,
  Delete,
  Get,
  Headers,
  HttpCode,
  Optional,
  Param,
  Patch,
  Post,
  Query,
  Res,
  UseGuards,
} from '@nestjs/common';
import { FastifyReply } from 'fastify';
import { randomUUID } from 'node:crypto';
import {
  runWithLlmUsageContext,
  type LlmUsageContext,
} from '../llm-usage/llm-usage.context';
import { LlmUsageService } from '../llm-usage/llm-usage.service';
import { TRACE_ID_HEADER } from '../common/constants/trace-id.constant';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { AuthGuard, AuthUserSession, Session } from '../auth/auth.guard';
import { ConversationRequestDto } from './dto/conversation-request.dto';
import { ListConversationsQueryDto } from './dto/list-conversations-query.dto';
import { RenameConversationDto } from './dto/rename-conversation.dto';
import { ToolResultDto } from './dto/tool-result.dto';
import { ContinueRunDto } from './dto/continue-run.dto';
import { ConversationService } from './services/conversation.service';
import { UsageBillingService } from '../credit/usage-billing.service';
import { ConcurrencyLimitService } from '../common/guards/concurrency-limit.service';
import { AI_USAGE_ACTION_TYPE } from '../credit/types/credit.types';
import { endSseResponse, initSseResponse, isSseResponse, writeSseEvent } from './utils/sse.util';

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
  constructor(
    private readonly conversationService: ConversationService,
    private readonly usageBilling: UsageBillingService,
    private readonly concurrencyLimit: ConcurrencyLimitService,
    @Optional() private readonly llmUsage?: LlmUsageService,
  ) {}

  @Post('conversation')
  @SkipEnvelope()
  async conversation(
    @Body() body: ConversationRequestDto,
    @Headers(TRACE_ID_HEADER) traceId: string | undefined,
    @Res() reply: FastifyReply,
    @Session() session: AuthUserSession | undefined,
  ): Promise<void> {
    // The trace id doubles as the prompt id, so a stepwise run (which stores
    // it) can attribute its later /continue waves back to this prompt.
    const promptId = traceId?.trim() && traceId.trim() !== '-' ? traceId.trim() : randomUUID();
    const context: LlmUsageContext = {
      promptId,
      userId: session?.user?.id,
      conversationId: body.conversationId,
    };
    this.llmUsage?.beginPrompt(context, {
      prompt: body.message,
      mode: body.mode,
      workbookId: body.workbookId,
    });
    await this.trackRequest(context, reply, () =>
      this.conversationService.handleConversation(body, reply, promptId, session?.user?.id),
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
    // Unknown until continueRun loads the run and re-points this at the run's
    // original prompt id; until then nothing is attributed to any prompt.
    const context: LlmUsageContext = { promptId: '', userId: session?.user?.id };
    await this.trackRequest(context, reply, () =>
      this.conversationService.continueRun(body, reply, traceId, session?.user?.id),
    );
  }

  private async trackRequest(
    context: LlmUsageContext,
    reply: FastifyReply,
    fn: () => Promise<void>,
  ): Promise<void> {
    const startedAt = Date.now();
    let error: string | undefined;

    if (context.userId && !(await this.usageBilling.canStart(context.userId))) {
      initSseResponse(reply);
      writeSseEvent(reply, 'error', {
        message: 'You are out of credits.',
        code: 'INSUFFICIENT_CREDIT',
        availableBalance: 0,
      });
      endSseResponse(reply);
      this.llmUsage?.endRequest(context, { durationMs: Date.now() - startedAt, error: 'insufficient_credit' });
      return;
    }

    // TASKS.md #344 — caps how many of THIS user's requests can be in flight
    // at once, independent of credit balance (a well-funded account could
    // otherwise fire unlimited concurrent LLM calls). Checked after the
    // credit gate so an already-blocked user sees the credit error, not a
    // generic rate-limit one.
    const acquired = context.userId ? this.concurrencyLimit.tryAcquire(context.userId) : true;
    if (!acquired) {
      initSseResponse(reply);
      writeSseEvent(reply, 'error', {
        message: 'You already have a request in progress. Wait for it to finish before sending another.',
        code: 'TOO_MANY_CONCURRENT_REQUESTS',
      });
      endSseResponse(reply);
      this.llmUsage?.endRequest(context, { durationMs: Date.now() - startedAt, error: 'too_many_concurrent_requests' });
      return;
    }

    this.settleBeforeStreamEnds(context, reply);
    try {
      await runWithLlmUsageContext(context, fn);
    } catch (err: unknown) {
      error = err instanceof Error ? err.message : String(err);
      throw err;
    } finally {
      if (context.userId) this.concurrencyLimit.release(context.userId);
      // Catches cost from calls that finished after the stream closed (e.g. a
      // Stop mid-wave); a no-op when the stream-end settle already billed all of it.
      void this.usageBilling.settle(context);
      this.llmUsage?.endRequest(context, { durationMs: Date.now() - startedAt, error });
    }
  }

  /**
   * ConversationService closes the SSE stream from ~70 places. Deferring the
   * real `end()` until this request's usage is debited lets the resulting
   * `credits` event reach the task pane, so its balance updates live.
   */
  private settleBeforeStreamEnds(context: LlmUsageContext, reply: FastifyReply): void {
    const billing = this.usageBilling;
    if (!context.userId) return;
    const raw = reply.raw;
    const end = raw.end.bind(raw) as (...args: unknown[]) => unknown;
    let ending = false;
    raw.end = ((...args: unknown[]) => {
      if (ending) return raw;
      ending = true;
      void billing
        .settle(context)
        .then((settlement) => {
          if (!settlement?.balances || raw.writableEnded || raw.destroyed || !isSseResponse(reply)) return;
          writeSseEvent(reply, 'credits', {
            ...settlement.balances,
            debited: settlement.debited,
            actionType: AI_USAGE_ACTION_TYPE,
            ...(context.conversationId ? { conversationId: context.conversationId } : {}),
          });
        })
        .catch(() => undefined)
        .finally(() => end(...args));
      return raw;
    }) as typeof raw.end;
  }
}
