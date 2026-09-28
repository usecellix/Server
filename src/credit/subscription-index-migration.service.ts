import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { InjectConnection } from '@nestjs/mongoose';
import { Connection } from 'mongoose';

/**
 * Drops the stale `stripeSubscriptionId_1` unique index left behind by the
 * Stripe->Razorpay migration (TASKS.md #181, credit-system-v2 session). The
 * `Subscription` schema (subscription.schema.ts) dropped the
 * `stripeSubscriptionId` field entirely in favor of `razorpaySubscriptionId`,
 * but Mongoose never drops indexes that a schema class no longer declares —
 * that requires an explicit migration. Left in place, EVERY subscription
 * document has `stripeSubscriptionId` absent (stored as null under this
 * index), so the unique constraint let exactly one upsert into `subscriptions`
 * succeed and made every subsequent one — the second Razorpay subscriber,
 * or any renewal event on `subscription.charged` — throw E11000 duplicate
 * key. RazorpayWebhookService.handleSubscriptionCredited has no try/catch, so
 * that exception propagated all the way to a bare 500 with no distinguishing
 * detail (HttpExceptionFilter reduces any non-HttpException to
 * "Unexpected error occurred"), which is what made every real purchase look
 * like "the webhook is broken" with no visible cause. Reproduced locally by
 * POSTing a correctly-HMAC-signed test payload straight at
 * /webhooks/razorpay and reading the (now-logged, see billing.controller.ts)
 * stack trace.
 *
 * Same OnModuleInit self-healing pattern as LogTtlIndexService — safe to run
 * on every boot (a no-op once the index is gone), so no manual `mongosh`
 * step is needed in any environment (a teammate's local DB, staging,
 * production) that still carries the old index.
 */
@Injectable()
export class SubscriptionIndexMigrationService implements OnModuleInit {
  private readonly logger = new Logger(SubscriptionIndexMigrationService.name);

  constructor(@InjectConnection() private readonly connection: Connection) {}

  async onModuleInit(): Promise<void> {
    try {
      const col = this.connection.collection('subscriptions');
      const indexes = await col.indexes();
      const staleIndex = indexes.find((idx) => idx.name === 'stripeSubscriptionId_1');
      if (staleIndex) {
        this.logger.warn(
          'Dropping stale stripeSubscriptionId_1 unique index (pre-Razorpay-migration leftover) — ' +
            'left in place, it E11000-collides on every subscription past the first.',
        );
        await col.dropIndex('stripeSubscriptionId_1');
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.logger.warn(`Failed to check/drop stripeSubscriptionId_1 index: ${msg}`);
    }
  }
}
