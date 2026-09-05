import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { CREDIT_COST_CATALOG, resolveCreditCost } from './credit-cost-catalog';
import { CreditActionType, CreditGateResult } from './types/credit.types';

/** Free tier's one-time grant, issued the first time a user is seen. CD-8. */
export const FREE_TIER_ONE_TIME_CREDITS = 30;

/**
 * Pre-flight balance check, run before any LLM call is dispatched.
 * CREDIT_SYSTEM.md CD-4 (no overdraft: finish the current task, block the
 * next request) and CD-6 (this is a read for the *next* request's gate, not
 * the atomic debit itself — the debit's own floor check in
 * CreditLedgerService.debit is what's authoritative under a race).
 */
@Injectable()
export class CreditGateService {
  constructor(
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
  ) {}

  /**
   * Checks whether billingEntityId can afford `actionType` at `quantity`
   * units. Returns allowed: false with a reason rather than throwing — the
   * caller (route/controller layer) decides how to surface a block to the
   * user. Lazily provisions a Free-tier account on first sight
   * (CREDIT_SYSTEM_SCHEMA.md §7's "new signups get one at signup instead of
   * backfilled" — this codebase has no signup hook yet, so first-gate-check
   * plays that role instead) rather than requiring a separate backfill.
   */
  async checkBalance(
    billingEntityId: string,
    actionType: CreditActionType,
    quantity = 1,
  ): Promise<CreditGateResult> {
    const catalogEntry = CREDIT_COST_CATALOG[actionType];
    if (!catalogEntry) {
      return { allowed: false, reason: 'unknown_action_type' };
    }
    if (catalogEntry.status === 'unwired') {
      return { allowed: false, reason: 'unwired_action' };
    }

    const requiredCredits = resolveCreditCost(actionType, quantity);
    const account = await this.ensureAccount(billingEntityId);

    const availableBalance = account.planCredits + account.purchasedCredits + account.oneTimeCredits;

    if (availableBalance < requiredCredits) {
      return { allowed: false, reason: 'insufficient_balance', availableBalance, requiredCredits };
    }
    return { allowed: true, availableBalance, requiredCredits };
  }

  /**
   * The CD-4 minimum-balance check for a Tier 3 dispatch, where the exact
   * cost isn't known until the run completes. Confirms the account has
   * *some* positive balance, not a specific amount — the completion-time
   * debit in CreditLedgerService.debit is what enforces the real floor.
   */
  async hasAnyBalance(billingEntityId: string): Promise<boolean> {
    const account = await this.ensureAccount(billingEntityId);
    return account.planCredits + account.purchasedCredits + account.oneTimeCredits > 0;
  }

  /**
   * Returns the account for billingEntityId, creating a Free-tier account
   * (30 oneTimeCredits, CD-8) the first time it's seen. `upsert` with
   * `setOnInsert` makes this safe under a race between two concurrent
   * first-requests for the same user — only one insert can win, and the
   * other's update becomes a no-op read of the winner's document.
   *
   * Public (not just used by checkBalance/hasAnyBalance) so
   * StripeCheckoutService/StripeWebhookService can guarantee a
   * credit_accounts document exists before attaching a Stripe customer id or
   * granting plan credits — a user who goes straight to checkout without
   * ever making a billable request first would otherwise have no account to
   * grant onto.
   */
  async ensureAccount(
    billingEntityId: string,
  ): Promise<Pick<CreditAccountDocument, 'planCredits' | 'purchasedCredits' | 'oneTimeCredits'>> {
    const account = await this.creditAccountModel
      .findOneAndUpdate(
        { billingEntityId },
        {
          $setOnInsert: {
            billingEntityType: 'user',
            billingEntityId,
            planTier: 'free',
            planCredits: 0,
            purchasedCredits: 0,
            oneTimeCredits: FREE_TIER_ONE_TIME_CREDITS,
          },
        },
        { upsert: true, new: true, projection: { planCredits: 1, purchasedCredits: 1, oneTimeCredits: 1 } },
      )
      .lean();
    return account!;
  }
}
