import { Body, Controller, HttpException, HttpStatus, Post, UseGuards } from '@nestjs/common';
import { AuthGuard, AuthUserSession, Session } from '../auth/auth.guard';
import { SkipEnvelope } from '../common/decorators/skip-envelope.decorator';
import { ConcurrencyLimitService } from '../common/guards/concurrency-limit.service';
import { InsufficientCreditError } from '../credit/errors/insufficient-credit.error';
import { WebChatAskDto, WebChatEstimateDto } from './dto/web-chat-ask.dto';
import { WebChatService } from './web-chat.service';

/**
 * Ask-mode chat for the `web/` Next.js app, answering over the user's stored
 * Excel add-in sessions.
 *
 * `userId` comes from the session, never the body — same rule
 * BillingController and ConversationController both state: accepting it as
 * input would let any caller read another user's history and spend their
 * credits.
 *
 * Listing/opening/deleting conversations is NOT duplicated here. Those routes
 * already exist on ConversationController (`GET /excel-ai/conversations`,
 * `GET|PATCH|DELETE /excel-ai/conversation/:id`) with the same AuthGuard and
 * the same ownership checks, and the web app calls them directly. A parallel
 * set of web-only CRUD routes would be a second place for that ownership
 * logic to drift.
 */
@UseGuards(AuthGuard)
@Controller('web-chat')
export class WebChatController {
  constructor(
    private readonly webChatService: WebChatService,
    private readonly concurrencyLimit: ConcurrencyLimitService,
  ) {}

  /**
   * Pre-send price for the composer's "Using N credits" hint. Separate from
   * `ask` (rather than a field on its response) because the UI needs the
   * number BEFORE the user commits — showing it only afterwards would defeat
   * CREDIT_SYSTEM.md CD-1's "know in advance what an action costs".
   */
  @Post('estimate')
  @SkipEnvelope()
  estimate(@Body() body: WebChatEstimateDto) {
    return this.webChatService.estimateCost(body.question, Boolean(body.conversationId));
  }

  @Post('ask')
  @SkipEnvelope()
  async ask(@Session() session: AuthUserSession, @Body() body: WebChatAskDto) {
    const userId = session.user.id;
    // TASKS.md #344 — same per-user in-flight cap the Excel add-in's
    // ConversationController applies, sharing ONE counter across both
    // surfaces (ConcurrencyLimitModule) since both spend from the same
    // credit account.
    if (!this.concurrencyLimit.tryAcquire(userId)) {
      throw new HttpException(
        {
          code: 'TOO_MANY_CONCURRENT_REQUESTS',
          message: 'You already have a request in progress. Wait for it to finish before sending another.',
        },
        HttpStatus.TOO_MANY_REQUESTS,
      );
    }
    try {
      return await this.webChatService.ask(userId, body.question, {
        conversationId: body.conversationId,
      });
    } catch (error) {
      if (error instanceof InsufficientCreditError) {
        // 402 rather than 400/403: the request is well-formed and the caller is
        // authenticated — what's missing is credit. The web client keys its
        // "Insufficient credits" state off this status specifically.
        throw new HttpException(
          {
            code: 'INSUFFICIENT_CREDIT',
            message: 'Insufficient credits. Upgrade your plan or buy a top-up pack.',
            requiredCredits: error.requiredCredits,
          },
          HttpStatus.PAYMENT_REQUIRED,
        );
      }
      throw error;
    } finally {
      this.concurrencyLimit.release(userId);
    }
  }
}
