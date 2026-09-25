import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntryDocument } from './schemas/credit-ledger.schema';
import { CREDIT_COST_CATALOG, resolveCreditCost } from './credit-cost-catalog';
import {
  AI_USAGE_ACTION_TYPE,
  CreditActionType,
  CreditBalances,
  CreditBucket,
  CreditDebitResult,
  UsageDebitResult,
} from './types/credit.types';
import { InsufficientCreditError } from './errors/insufficient-credit.error';

interface DebitContext {
  conversationId?: string;
  changeSetId?: string;
  /** WHO spent it, for pooled org accounts. CREDIT_SYSTEM.md CD-9. */
  seatUserId?: string;
}

/**
 * Owns the two operations that change a balance: the completion-time debit
 * (CREDIT_SYSTEM.md CD-3) and grants/purchases (the Razorpay webhook
 * call-ins — grantPlanCredits / addPurchasedCredits). Every balance change is
 * paired with an append-only credit_ledger row (CD-9).
 */
@Injectable()
export class CreditLedgerService {
  constructor(
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
    @InjectModel(CreditLedgerEntry.name)
    private readonly creditLedgerModel: Model<CreditLedgerEntryDocument>,
  ) {}

  /**
   * Atomically debits `actionType` at `quantity` units from billingEntityId,
   * consuming planCredits, then purchasedCredits, then oneTimeCredits
   * (CD-8's fixed order), inside one findOneAndUpdate so the floor check and
   * the fallthrough math can't race against a concurrent debit (CD-6).
   * Never throws for a plain insufficient-balance outcome — returns
   * `{ debited: false }` — so a caller that already gated via
   * CreditGateService can treat a race-lost debit as a normal, expected
   * outcome rather than an exceptional one.
   */
  async debit(
    billingEntityId: string,
    actionType: CreditActionType,
    quantity = 1,
    context: DebitContext = {},
  ): Promise<CreditDebitResult> {
    const catalogEntry = CREDIT_COST_CATALOG[actionType];
    if (!catalogEntry || catalogEntry.status === 'unwired') {
      throw new InsufficientCreditError(billingEntityId, 0);
    }
    const cost = resolveCreditCost(actionType, quantity);

    // Read-before-write, but not the authority on whether the debit succeeds
    // — that's the atomic update's own $expr floor check below (CD-6). This
    // read only supplies the ledger with per-bucket consumption amounts; if
    // it races with a concurrent debit, the bucket split recorded may be
    // approximate, but the atomic update's total is always correct and the
    // ledger's sum(amount) across a debit's rows always equals -cost.
    const preDebit = await this.creditAccountModel
      .findOne({ billingEntityId }, { planCredits: 1, purchasedCredits: 1, oneTimeCredits: 1 })
      .lean();

    const updated = await this.creditAccountModel.findOneAndUpdate(
      {
        billingEntityId,
        $expr: { $gte: [{ $add: ['$planCredits', '$purchasedCredits', '$oneTimeCredits'] }, cost] },
      },
      [
        {
          $set: {
            planCredits: { $max: [0, { $subtract: ['$planCredits', cost] }] },
            purchasedCredits: {
              $max: [
                0,
                {
                  $subtract: [
                    '$purchasedCredits',
                    { $max: [0, { $subtract: [cost, '$planCredits'] }] },
                  ],
                },
              ],
            },
            oneTimeCredits: {
              $max: [
                0,
                {
                  $subtract: [
                    '$oneTimeCredits',
                    {
                      $max: [
                        0,
                        { $subtract: [cost, { $add: ['$planCredits', '$purchasedCredits'] }] },
                      ],
                    },
                  ],
                },
              ],
            },
          },
        },
      ],
      // The update above is an aggregation PIPELINE (an array), which is what
      // makes the floor check and the fixed consumption order atomic (CD-6 /
      // CD-8). Mongoose 9 refuses an array update unless the caller says it
      // means a pipeline — without this flag every billable request throws
      // "Cannot pass an array to query updates unless the `updatePipeline`
      // option is set" at debit time, which surfaced in the task pane as a red
      // error on a request the pipeline had already planned and verified.
      // The unit specs mock findOneAndUpdate, so only a live Mongo call shows
      // it. TASKS.md #237.
      { new: true, updatePipeline: true },
    );

    if (!updated) {
      return { debited: false };
    }

    await this.recordDebitLedgerRows(
      billingEntityId,
      actionType,
      cost,
      preDebit ?? { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 },
      context,
    );

    return {
      debited: true,
      balances: {
        planCredits: updated.planCredits,
        purchasedCredits: updated.purchasedCredits,
        oneTimeCredits: updated.oneTimeCredits,
      },
    };
  }

  /**
   * Debits up to `credits` for real AI usage. Unlike `debit`, a balance
   * smaller than the charge is not refused: whatever is left is taken, so the
   * balance reaches exactly 0 instead of freezing at a value below the next
   * charge. Atomic: the clamp and the plan → purchased → one-time consumption
   * order run inside one pipeline update, and the pre-update document gives
   * the exact per-bucket split for the ledger.
   */
  async debitUsage(
    billingEntityId: string,
    credits: number,
    context: DebitContext & { promptId?: string; costUsd?: number } = {},
  ): Promise<UsageDebitResult> {
    const requested = Math.max(0, Math.ceil(credits));
    if (requested === 0) return { debited: 0, requested };

    const total = { $add: ['$planCredits', '$purchasedCredits', '$oneTimeCredits'] };
    const take = '$__usageDebit';
    const before = await this.creditAccountModel
      .findOneAndUpdate(
        { billingEntityId, $expr: { $gt: [total, 0] } },
        [
          { $set: { __usageDebit: { $min: [requested, total] } } },
          {
            $set: {
              planCredits: { $max: [0, { $subtract: ['$planCredits', take] }] },
              purchasedCredits: {
                $max: [
                  0,
                  { $subtract: ['$purchasedCredits', { $max: [0, { $subtract: [take, '$planCredits'] }] }] },
                ],
              },
              oneTimeCredits: {
                $max: [
                  0,
                  {
                    $subtract: [
                      '$oneTimeCredits',
                      { $max: [0, { $subtract: [take, { $add: ['$planCredits', '$purchasedCredits'] }] }] },
                    ],
                  },
                ],
              },
            },
          },
          { $unset: '__usageDebit' },
        ],
        // Pipeline update needs updatePipeline (TASKS.md #237); the pre-debit
        // document is what the exact per-bucket split below is computed from.
        {
          returnDocument: 'before',
          updatePipeline: true,
          projection: { planCredits: 1, purchasedCredits: 1, oneTimeCredits: 1 },
        },
      )
      .lean();

    if (!before) return { debited: 0, requested };

    const split = splitAcrossBuckets(requested, before);
    const debited = split.reduce((sum, row) => sum + row.amount, 0);
    const balances: CreditBalances = {
      planCredits: before.planCredits,
      purchasedCredits: before.purchasedCredits,
      oneTimeCredits: before.oneTimeCredits,
    };
    for (const row of split) balances[row.bucket] -= row.amount;

    await this.creditLedgerModel.insertMany(
      split.map((row, index) => ({
        billingEntityId,
        seatUserId: context.seatUserId,
        entryType: 'debit' as const,
        amount: -row.amount,
        bucket: row.bucket,
        actionType: AI_USAGE_ACTION_TYPE,
        conversationId: context.conversationId,
        changeSetId: context.changeSetId,
        promptId: context.promptId,
        ...(index === 0 && context.costUsd !== undefined ? { costUsd: context.costUsd } : {}),
        createdAt: new Date(),
      })),
    );

    return { debited, requested, balances };
  }

  /**
   * Monthly plan allotment grant — Razorpay `subscription.activated`/
   * `subscription.charged` call-in. RESETS `planCredits` to `amount` rather
   * than adding to it (TASKS.md #342): `credit-account.schema.ts` documents
   * planCredits as resetting each cycle, not rolling over, but the previous
   * `$inc` stacked every renewal onto whatever was left, so a subscriber who
   * under-used one month kept compounding a balance no plan was ever priced
   * to cover. Any credits forfeited by the reset are logged as a separate
   * `expire` row so the ledger shows where they went, rather than a `grant`
   * row silently overstating what the renewal actually added.
   */
  async grantPlanCredits(billingEntityId: string, amount: number, paymentEventId?: string): Promise<void> {
    const before = await this.creditAccountModel
      .findOneAndUpdate(
        { billingEntityId },
        { $set: { planCredits: amount } },
        { new: false, upsert: true, projection: { planCredits: 1 } },
      )
      .lean();
    const forfeited = Math.max(0, before?.planCredits ?? 0);

    const rows: Record<string, unknown>[] = [
      // paymentEventId is this row's idempotency key (unique index) — the
      // real Razorpay event id belongs on exactly one row, not both.
      { billingEntityId, entryType: 'grant', amount, bucket: 'planCredits', paymentEventId, createdAt: new Date() },
    ];
    if (forfeited > 0) {
      rows.push({
        billingEntityId,
        entryType: 'expire',
        amount: -forfeited,
        bucket: 'planCredits',
        createdAt: new Date(),
      });
    }
    await this.creditLedgerModel.insertMany(rows);
  }

  /** Top-up pack purchase — Razorpay `payment_link.paid` call-in. */
  async addPurchasedCredits(billingEntityId: string, amount: number, paymentEventId?: string): Promise<void> {
    await this.creditAccountModel.updateOne(
      { billingEntityId },
      { $inc: { purchasedCredits: amount } },
    );
    await this.creditLedgerModel.create({
      billingEntityId,
      entryType: 'purchase',
      amount,
      bucket: 'purchasedCredits',
      paymentEventId,
      createdAt: new Date(),
    });
  }

  /**
   * Splits `cost` across buckets in CD-8's fixed order (planCredits, then
   * purchasedCredits, then oneTimeCredits) against the pre-debit balances,
   * and writes one append-only ledger row per bucket actually touched — a
   * single row would hide which buckets a debit drew from, which CD-9 needs
   * for support/dispute resolution and per-seat usage analytics.
   */
  private async recordDebitLedgerRows(
    billingEntityId: string,
    actionType: CreditActionType,
    cost: number,
    preDebit: Pick<CreditAccountDocument, 'planCredits' | 'purchasedCredits' | 'oneTimeCredits'>,
    context: DebitContext,
  ): Promise<void> {
    const rows = splitAcrossBuckets(cost, preDebit);

    await this.creditLedgerModel.insertMany(
      rows.map((row) => ({
        billingEntityId,
        seatUserId: context.seatUserId,
        entryType: 'debit' as const,
        amount: -row.amount,
        bucket: row.bucket,
        actionType,
        conversationId: context.conversationId,
        changeSetId: context.changeSetId,
        createdAt: new Date(),
      })),
    );
  }
}

/** CD-8's fixed consumption order: planCredits, then purchasedCredits, then oneTimeCredits. */
function splitAcrossBuckets(
  cost: number,
  balances: Pick<CreditBalances, CreditBucket>,
): { bucket: CreditBucket; amount: number }[] {
  let remaining = cost;
  const rows: { bucket: CreditBucket; amount: number }[] = [];
  for (const bucket of ['planCredits', 'purchasedCredits', 'oneTimeCredits'] as const) {
    if (remaining <= 0) break;
    const consumed = Math.min(remaining, Math.max(0, balances[bucket]));
    if (consumed > 0) {
      rows.push({ bucket, amount: consumed });
      remaining -= consumed;
    }
  }
  return rows;
}
