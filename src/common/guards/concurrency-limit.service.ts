import { Injectable } from '@nestjs/common';
import { AppConfigService } from '../../config/app-config.service';

/**
 * Caps how many LLM-spending requests one signed-in user can have in flight
 * at once (TASKS.md #344). Before this, a signed-in user could fire unlimited
 * concurrent requests against `/excel-ai/conversation`/`/continue` or
 * `/web-chat/ask` — each one dispatches at least one real OpenRouter call, so
 * nothing stood between a stuck/duplicated client (or a scripted abuser) and
 * unbounded parallel spend.
 *
 * Deliberately an in-memory per-process counter, not a distributed limiter
 * (Redis, Mongo doc, etc.): this deployment runs a single backend process
 * (`ARCHITECTURE.md` §6, "no message queue / job runner exists"), so a
 * process-local Map is exactly as correct as a distributed store would be,
 * at zero added infrastructure. Revisit if the backend is ever horizontally
 * scaled — a per-instance cap would then undercount a user's true
 * concurrency by a factor of instance count.
 *
 * Guards admission (`CanActivate`) can't own this cleanly by themselves: a
 * guard's `canActivate()` can refuse a request, but there is no matching
 * "the request is done" hook a plain guard receives, and Fastify's own
 * request lifecycle finishes before an SSE response's stream actually
 * closes. `tryAcquire`/`release` are instead called directly around the
 * request body (`ConversationController.trackRequest`,
 * `WebChatController.ask`), which already spans exactly that lifetime for
 * the usage-billing settle logic (TASKS.md #341) — the same seam, reused.
 */
@Injectable()
export class ConcurrencyLimitService {
  private readonly inFlight = new Map<string, number>();

  constructor(private readonly config: AppConfigService) {}

  /** True and increments the counter if under the cap; false (no state change) otherwise. */
  tryAcquire(userId: string): boolean {
    const current = this.inFlight.get(userId) ?? 0;
    if (current >= this.config.maxConcurrentRequestsPerUser) return false;
    this.inFlight.set(userId, current + 1);
    return true;
  }

  /**
   * Idempotent-safe against being called without a matching successful
   * `tryAcquire` (e.g. a caller that releases unconditionally in a `finally`
   * after an early return) — never goes below 0, never throws on an unknown
   * user.
   */
  release(userId: string): void {
    const current = this.inFlight.get(userId) ?? 0;
    if (current <= 1) {
      this.inFlight.delete(userId);
    } else {
      this.inFlight.set(userId, current - 1);
    }
  }

  /** Test/observability hook — not used on the request path. */
  currentCount(userId: string): number {
    return this.inFlight.get(userId) ?? 0;
  }
}
