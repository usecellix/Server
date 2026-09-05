import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntryDocument } from './schemas/credit-ledger.schema';
import { CREDIT_COST_CATALOG, resolveCreditCost } from './credit-cost-catalog';
import { CreditActionType, CreditBucket, CreditDebitResult } from './types/credit.types';
import { InsufficientCreditError } from './errors/insufficient-credit.error';

interface DebitContext {
  conversationId?: string;
  changeSetId?: string;
  /** WHO spent it, for pooled org accounts. CREDIT_SYSTEM.md CD-9. */
  seatUserId?: string;
}

/**
 * Owns the two operations that change a balance: the completion-time debit
 * (CREDIT_SYSTEM.md CD-3) and grants/purchases (the Stripe webhook call-ins
 * this PR stubs per §7 — grantPlanCredits / addPurchasedCredits). Every
 * balance change is paired with an append-only credit_ledger row (CD-9).
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
      { new: true },
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

  /** Monthly plan allotment grant — Stripe `invoice.paid` / renewal call-in (§7). */
  async grantPlanCredits(billingEntityId: string, amount: number, stripeEventId?: string): Promise<void> {
    await this.creditAccountModel.updateOne(
      { billingEntityId },
      { $inc: { planCredits: amount } },
    );
    await this.creditLedgerModel.create({
      billingEntityId,
      entryType: 'grant',
      amount,
      bucket: 'planCredits',
      stripeEventId,
      createdAt: new Date(),
    });
  }

  /** Top-up pack purchase — Stripe `checkout.session.completed` call-in (§7). */
  async addPurchasedCredits(billingEntityId: string, amount: number, stripeEventId?: string): Promise<void> {
    await this.creditAccountModel.updateOne(
      { billingEntityId },
      { $inc: { purchasedCredits: amount } },
    );
    await this.creditLedgerModel.create({
      billingEntityId,
      entryType: 'purchase',
      amount,
      bucket: 'purchasedCredits',
      stripeEventId,
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
    let remaining = cost;
    const rows: { bucket: CreditBucket; amount: number }[] = [];

    for (const bucket of ['planCredits', 'purchasedCredits', 'oneTimeCredits'] as const) {
      if (remaining <= 0) break;
      const consumed = Math.min(remaining, Math.max(0, preDebit[bucket]));
      if (consumed > 0) {
        rows.push({ bucket, amount: consumed });
        remaining -= consumed;
      }
    }

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
