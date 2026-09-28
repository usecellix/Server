import { ConversationController } from '../src/excel-ai/conversation.controller';
import { ConcurrencyLimitService } from '../src/common/guards/concurrency-limit.service';
import type { AuthUserSession } from '../src/auth/auth.guard';
import type { FastifyReply } from 'fastify';

/**
 * Minimal SSE-capable FastifyReply stand-in: real enough for
 * initSseResponse/writeSseEvent/endSseResponse (utils/sse.util.ts) to run
 * against without throwing, so `trackRequest`'s early-return paths (credit
 * gate, concurrency gate) can be exercised end to end through the real
 * controller method rather than only unit-testing the pieces in isolation.
 */
function fakeReply(): { reply: FastifyReply; writes: string[]; ended: boolean } {
  const writes: string[] = [];
  const state = { ended: false };
  const raw = {
    writeHead: jest.fn(),
    write: jest.fn((chunk: string) => {
      writes.push(chunk);
      return true;
    }),
    end: jest.fn(() => {
      state.ended = true;
    }),
    writableEnded: false,
    destroyed: false,
    once: jest.fn(),
  };
  return { reply: { raw } as unknown as FastifyReply, writes, get ended() { return state.ended; } } as never;
}

function session(userId: string): AuthUserSession {
  return { user: { id: userId, email: 'ca@example.com' } } as unknown as AuthUserSession;
}

describe('ConversationController — per-user concurrency cap (TASKS.md #344)', () => {
  it('refuses a second concurrent request from the same user with a TOO_MANY_CONCURRENT_REQUESTS SSE error, without calling the service', async () => {
    const concurrencyLimit = new ConcurrencyLimitService({ maxConcurrentRequestsPerUser: 1 } as never);
    // Occupy the one slot directly, simulating a first request already in flight.
    concurrencyLimit.tryAcquire('user-1');

    const handleConversation = jest.fn().mockResolvedValue(undefined);
    const usageBilling = { canStart: jest.fn().mockResolvedValue(true), settle: jest.fn().mockResolvedValue(null) };
    const controller = new ConversationController(
      { handleConversation } as never,
      usageBilling as never,
      concurrencyLimit,
    );

    const { reply, writes } = fakeReply();
    await controller.conversation({ message: 'hi' } as never, '-', reply, session('user-1'));

    expect(handleConversation).not.toHaveBeenCalled();
    expect(writes.some((w) => w.includes('TOO_MANY_CONCURRENT_REQUESTS'))).toBe(true);
    // The slot this rejected request never held must not have been consumed —
    // the original in-flight request's slot is still the only one occupied.
    expect(concurrencyLimit.currentCount('user-1')).toBe(1);
  });

  it('admits a request when under the cap, and releases the slot once the handler settles', async () => {
    const concurrencyLimit = new ConcurrencyLimitService({ maxConcurrentRequestsPerUser: 1 } as never);
    const handleConversation = jest.fn().mockResolvedValue(undefined);
    const usageBilling = { canStart: jest.fn().mockResolvedValue(true), settle: jest.fn().mockResolvedValue(null) };
    const controller = new ConversationController(
      { handleConversation } as never,
      usageBilling as never,
      concurrencyLimit,
    );

    const { reply } = fakeReply();
    await controller.conversation({ message: 'hi' } as never, '-', reply, session('user-1'));

    expect(handleConversation).toHaveBeenCalledTimes(1);
    expect(concurrencyLimit.currentCount('user-1')).toBe(0);
  });

  it('releases the slot even when the underlying handler throws', async () => {
    const concurrencyLimit = new ConcurrencyLimitService({ maxConcurrentRequestsPerUser: 1 } as never);
    const handleConversation = jest.fn().mockRejectedValue(new Error('boom'));
    const usageBilling = { canStart: jest.fn().mockResolvedValue(true), settle: jest.fn().mockResolvedValue(null) };
    const controller = new ConversationController(
      { handleConversation } as never,
      usageBilling as never,
      concurrencyLimit,
    );

    const { reply } = fakeReply();
    await expect(
      controller.conversation({ message: 'hi' } as never, '-', reply, session('user-1')),
    ).rejects.toThrow('boom');

    expect(concurrencyLimit.currentCount('user-1')).toBe(0);
  });

  it('checks the credit gate before the concurrency gate, so an out-of-credit user sees the credit error', async () => {
    const concurrencyLimit = new ConcurrencyLimitService({ maxConcurrentRequestsPerUser: 1 } as never);
    concurrencyLimit.tryAcquire('user-1'); // slot already occupied too — credit check must still win
    const handleConversation = jest.fn().mockResolvedValue(undefined);
    const usageBilling = { canStart: jest.fn().mockResolvedValue(false), settle: jest.fn().mockResolvedValue(null) };
    const controller = new ConversationController(
      { handleConversation } as never,
      usageBilling as never,
      concurrencyLimit,
    );

    const { reply, writes } = fakeReply();
    await controller.conversation({ message: 'hi' } as never, '-', reply, session('user-1'));

    expect(writes.some((w) => w.includes('INSUFFICIENT_CREDIT'))).toBe(true);
    expect(writes.some((w) => w.includes('TOO_MANY_CONCURRENT_REQUESTS'))).toBe(false);
    expect(handleConversation).not.toHaveBeenCalled();
  });
});
