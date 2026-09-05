import { BadRequestException, Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { CreditAccount, CreditAccountDocument } from './schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntryDocument } from './schemas/credit-ledger.schema';

const LEDGER_DEFAULT_LIMIT = 25;
const LEDGER_MAX_LIMIT = 100;

export interface AccountSummary {
  billingEntityType: 'user' | 'org';
  planTier: 'free' | 'solo' | 'firm' | 'enterprise';
  planCredits: number;
  purchasedCredits: number;
  oneTimeCredits: number;
  availableBalance: number;
  currentPeriodEnd: Date | null;
}

export interface LedgerPage {
  entries: Array<{
    entryType: 'grant' | 'purchase' | 'debit' | 'one_time_grant';
    amount: number;
    bucket: 'planCredits' | 'purchasedCredits' | 'oneTimeCredits';
    actionType?: string;
    seatUserId?: string;
    createdAt: Date;
  }>;
  nextCursor: string | null;
}

/**
 * Read-side for the `/billing/*` GET routes (CREDIT_SYSTEM_SCHEMA.md §5).
 * Deliberately separate from CreditGateService/CreditLedgerService, which
 * own writes — this service only reads, and (unlike the gate check) never
 * lazily provisions an account, since a balance/history page for an
 * account that has never made a billable request should show the honest
 * "no account yet" state rather than silently creating one.
 */
@Injectable()
export class CreditAccountQueryService {
  constructor(
    @InjectModel(CreditAccount.name)
    private readonly creditAccountModel: Model<CreditAccountDocument>,
    @InjectModel(CreditLedgerEntry.name)
    private readonly creditLedgerModel: Model<CreditLedgerEntryDocument>,
  ) {}

  async getAccountSummary(billingEntityId: string): Promise<AccountSummary | null> {
    const account = await this.creditAccountModel.findOne({ billingEntityId }).lean();
    if (!account) return null;
    return {
      billingEntityType: account.billingEntityType,
      planTier: account.planTier,
      planCredits: account.planCredits,
      purchasedCredits: account.purchasedCredits,
      oneTimeCredits: account.oneTimeCredits,
      availableBalance: account.planCredits + account.purchasedCredits + account.oneTimeCredits,
      currentPeriodEnd: account.currentPeriodEnd ?? null,
    };
  }

  async getLedgerPage(
    billingEntityId: string,
    options: { limit?: number; cursor?: string } = {},
  ): Promise<LedgerPage> {
    const limit = Math.min(Math.max(Math.trunc(options.limit ?? LEDGER_DEFAULT_LIMIT), 1), LEDGER_MAX_LIMIT);

    const filter: Record<string, unknown> = { billingEntityId };
    if (options.cursor) {
      const cursorDate = new Date(options.cursor);
      if (Number.isNaN(cursorDate.getTime())) {
        throw new BadRequestException('INVALID_CURSOR');
      }
      filter.createdAt = { $lt: cursorDate };
    }

    // limit + 1 so "is there another page" comes from this query, not a
    // second count() that could disagree with it — same pattern
    // ConversationService.listConversations uses.
    const docs = await this.creditLedgerModel
      .find(filter)
      .sort({ createdAt: -1 })
      .limit(limit + 1)
      .select('entryType amount bucket actionType seatUserId createdAt')
      .lean();

    const hasMore = docs.length > limit;
    const page = hasMore ? docs.slice(0, limit) : docs;

    return {
      entries: page.map((doc) => ({
        entryType: doc.entryType,
        amount: doc.amount,
        bucket: doc.bucket,
        actionType: doc.actionType,
        seatUserId: doc.seatUserId,
        createdAt: doc.createdAt,
      })),
      nextCursor: hasMore ? page[page.length - 1].createdAt.toISOString() : null,
    };
  }
}
