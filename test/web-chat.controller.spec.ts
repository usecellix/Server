import { HttpException, HttpStatus } from '@nestjs/common';
import { WebChatController } from '../src/web-chat/web-chat.controller';
import { InsufficientCreditError } from '../src/credit/errors/insufficient-credit.error';
import { ConcurrencyLimitService } from '../src/common/guards/concurrency-limit.service';
import { AuthUserSession } from '../src/auth/auth.guard';

function session(userId: string): AuthUserSession {
  return { user: { id: userId, email: 'ca@example.com' } } as unknown as AuthUserSession;
}

/** Always-admits stub — the concurrency cap itself is unit-tested on ConcurrencyLimitService directly. */
function allowingLimiter(): ConcurrencyLimitService {
  return { tryAcquire: jest.fn().mockReturnValue(true), release: jest.fn() } as unknown as ConcurrencyLimitService;
}

describe('WebChatController', () => {
  it('resolves userId from the session, never the request body', async () => {
    // Same rule BillingController and ConversationController state: accepting
    // a userId as input would let any caller read another user's history and
    // spend their credits.
    const ask = jest.fn().mockResolvedValue({ answer: 'ok' });
    const controller = new WebChatController({ ask } as never, allowingLimiter());

    await controller.ask(session('user-1'), {
      question: 'what did I do',
      conversationId: 'c1',
    });

    expect(ask).toHaveBeenCalledWith('user-1', 'what did I do', { conversationId: 'c1' });
  });

  it('translates an insufficient-credit error into a 402, not a 400 or 500', async () => {
    // The request is well-formed and the caller is authenticated — what's
    // missing is credit. The web client keys its upgrade prompt off this status.
    const ask = jest.fn().mockRejectedValue(new InsufficientCreditError('user-1', 5));
    const controller = new WebChatController({ ask } as never, allowingLimiter());

    await expect(
      controller.ask(session('user-1'), { question: 'compare everything' }),
    ).rejects.toMatchObject({ status: HttpStatus.PAYMENT_REQUIRED });
  });

  it('carries the required-credit figure through, so the client can say how much is needed', async () => {
    const ask = jest.fn().mockRejectedValue(new InsufficientCreditError('user-1', 5));
    const controller = new WebChatController({ ask } as never, allowingLimiter());

    try {
      await controller.ask(session('user-1'), { question: 'compare everything' });
      throw new Error('expected a 402');
    } catch (error) {
      expect(error).toBeInstanceOf(HttpException);
      expect((error as HttpException).getResponse()).toMatchObject({
        code: 'INSUFFICIENT_CREDIT',
        requiredCredits: 5,
      });
    }
  });

  it('lets a non-credit failure surface unchanged rather than mislabelling it as a billing problem', async () => {
    const ask = jest.fn().mockRejectedValue(new Error('provider down'));
    const controller = new WebChatController({ ask } as never, allowingLimiter());

    await expect(
      controller.ask(session('user-1'), { question: 'what did I do' }),
    ).rejects.toThrow('provider down');
  });

  it('estimates without touching the session — price depends on the question, not who asks', async () => {
    const estimateCost = jest.fn().mockReturnValue({ complexity: 'simple', credits: 2 });
    const controller = new WebChatController({ estimateCost } as never, allowingLimiter());

    const result = controller.estimate({ question: 'what did I do' });

    expect(estimateCost).toHaveBeenCalledWith('what did I do', false);
    expect(result).toEqual({ complexity: 'simple', credits: 2 });
  });

  it('tells the estimator when a question is scoped to one session', async () => {
    const estimateCost = jest.fn().mockReturnValue({ complexity: 'simple', credits: 2 });
    const controller = new WebChatController({ estimateCost } as never, allowingLimiter());

    controller.estimate({ question: 'what changed', conversationId: 'c1' });

    expect(estimateCost).toHaveBeenCalledWith('what changed', true);
  });

  describe('concurrency limit (TASKS.md #344)', () => {
    it('rejects with 429 before calling ask() when the user already has a request in flight', async () => {
      const ask = jest.fn().mockResolvedValue({ answer: 'ok' });
      const concurrencyLimit = { tryAcquire: jest.fn().mockReturnValue(false), release: jest.fn() } as unknown as ConcurrencyLimitService;
      const controller = new WebChatController({ ask } as never, concurrencyLimit);

      await expect(
        controller.ask(session('user-1'), { question: 'what did I do' }),
      ).rejects.toMatchObject({ status: HttpStatus.TOO_MANY_REQUESTS });
      expect(ask).not.toHaveBeenCalled();
    });

    it('releases the slot even when ask() throws, so one failed request does not permanently occupy it', async () => {
      const ask = jest.fn().mockRejectedValue(new Error('provider down'));
      const concurrencyLimit = { tryAcquire: jest.fn().mockReturnValue(true), release: jest.fn() } as unknown as ConcurrencyLimitService;
      const controller = new WebChatController({ ask } as never, concurrencyLimit);

      await expect(controller.ask(session('user-1'), { question: 'x' })).rejects.toThrow('provider down');

      expect(concurrencyLimit.release).toHaveBeenCalledWith('user-1');
    });

    it('releases the slot on a normal successful response', async () => {
      const ask = jest.fn().mockResolvedValue({ answer: 'ok' });
      const concurrencyLimit = { tryAcquire: jest.fn().mockReturnValue(true), release: jest.fn() } as unknown as ConcurrencyLimitService;
      const controller = new WebChatController({ ask } as never, concurrencyLimit);

      await controller.ask(session('user-1'), { question: 'x' });

      expect(concurrencyLimit.release).toHaveBeenCalledWith('user-1');
    });
  });
});
