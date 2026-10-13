import { Injectable } from '@nestjs/common';
import { InjectConnection, InjectModel } from '@nestjs/mongoose';
import { Connection, Model } from 'mongoose';
import { ObjectId } from 'mongodb';
import { CreditAccount, CreditAccountDocument } from '../credit/schemas/credit-account.schema';
import { CreditLedgerEntry, CreditLedgerEntryDocument } from '../credit/schemas/credit-ledger.schema';
import { Subscription, SubscriptionDocument } from '../credit/schemas/subscription.schema';
import { AdminPromptsService } from './admin-prompts.service';

const USERS_PAGE_SIZE = 25;
/** Matches the original Dashboard/src/lib/data/users.ts getUser's `.limit(15)` on ai_prompts. */
const RECENT_PROMPTS_LIMIT = 15;

export type UserSort = 'recent' | 'spend' | 'prompts' | 'credits' | 'creditsUsed' | 'seen';
const SORT_FIELD: Record<UserSort, string> = {
  recent: 'createdAt',
  spend: 'costUsd',
  prompts: 'prompts',
  credits: 'credits',
  creditsUsed: 'creditsUsed',
  seen: 'lastSeenAt',
};

export interface AdminUserRow {
  id: string;
  name: string;
  email: string;
  image: string | null;
  createdAt: string | null;
  lastSeenAt: string | null;
  plan: string;
  subscriptionStatus: string | null;
  credits: number | null;
  creditsUsed: number;
  prompts: number;
  costUsd: number;
  lastPromptAt: string | null;
}

export interface AdminUserListFilters {
  q?: string;
  page: number;
  sort: UserSort;
  plan?: string;
  since?: Date;
}

function num(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Same read shape `Dashboard/src/lib/data/users.ts` implemented directly
 * against Mongo — moved here so the Dashboard calls an authenticated backend
 * route instead of holding its own MongoClient. Kept intentionally close to
 * the original so the Dashboard UI needs no changes beyond swapping its data
 * layer's transport (TASKS.md admin-api-migration).
 */
@Injectable()
export class AdminUsersService {
  constructor(
    @InjectConnection() private readonly connection: Connection,
    @InjectModel(CreditAccount.name) private readonly creditAccountModel: Model<CreditAccountDocument>,
    @InjectModel(CreditLedgerEntry.name) private readonly creditLedgerModel: Model<CreditLedgerEntryDocument>,
    @InjectModel(Subscription.name) private readonly subscriptionModel: Model<SubscriptionDocument>,
    private readonly adminPrompts: AdminPromptsService,
  ) {}

  async listUsers(filters: AdminUserListFilters): Promise<{ rows: AdminUserRow[]; total: number }> {
    const db = this.connection.db;
    if (!db) return { rows: [], total: 0 };

    const where: Record<string, unknown> = {};
    if (filters.since) where.createdAt = { $gte: filters.since };
    if (filters.q) {
      const pattern = { $regex: escapeRegex(filters.q.slice(0, 200)), $options: 'i' };
      where.$or = [{ email: pattern }, { name: pattern }];
    }

    const [docs, total] = await Promise.all([
      db
        .collection('user')
        .find(where, { projection: { name: 1, email: 1, image: 1, createdAt: 1 } })
        .sort({ createdAt: -1 })
        .toArray(),
      db.collection('user').countDocuments(where),
    ]);
    if (docs.length === 0) return { rows: [], total: 0 };

    const ids = docs.map((d) => d._id.toHexString());
    const [accounts, subs, usage, sessions, used] = await Promise.all([
      this.creditAccountModel.find({ billingEntityId: { $in: ids } }).lean(),
      this.subscriptionModel.find({ billingEntityId: { $in: ids } }).sort({ updatedAt: -1 }).lean(),
      db
        .collection('ai_prompts')
        .aggregate<{ _id: string; prompts: number; costUsd: number; last: Date }>([
          { $match: { userId: { $in: ids } } },
          { $group: { _id: '$userId', prompts: { $sum: 1 }, costUsd: { $sum: '$costUsd' }, last: { $max: '$createdAt' } } },
        ])
        .toArray(),
      db
        .collection('session')
        .aggregate<{ _id: unknown; last: Date }>([
          { $match: { userId: { $in: docs.map((d) => d._id) } } },
          { $group: { _id: '$userId', last: { $max: '$updatedAt' } } },
        ])
        .toArray(),
      this.creditLedgerModel
        .aggregate<{ _id: string; credits: number }>([
          { $match: { billingEntityId: { $in: ids }, entryType: 'debit' } },
          { $group: { _id: '$billingEntityId', credits: { $sum: { $multiply: ['$amount', -1] } } } },
        ]),
    ]);

    const accountBy = new Map(accounts.map((a) => [a.billingEntityId, a]));
    const subBy = new Map<string, (typeof subs)[number]>();
    for (const sub of subs) if (!subBy.has(sub.billingEntityId)) subBy.set(sub.billingEntityId, sub);
    const usageBy = new Map(usage.map((u) => [u._id, u]));
    const seenBy = new Map(sessions.map((s) => [String(s._id), s.last]));
    const usedBy = new Map(used.map((u) => [u._id, num(u.credits)]));

    let rows: AdminUserRow[] = docs.map((doc) => {
      const id = doc._id.toHexString();
      const account = accountBy.get(id);
      const sub = subBy.get(id);
      const use = usageBy.get(id);
      const seen = seenBy.get(id);
      const total = account ? num(account.planCredits) + num(account.purchasedCredits) + num(account.oneTimeCredits) : null;
      return {
        id,
        name: doc.name ?? '',
        email: doc.email ?? '',
        image: doc.image ?? null,
        createdAt: doc.createdAt ? new Date(doc.createdAt).toISOString() : null,
        lastSeenAt: seen ? new Date(seen).toISOString() : null,
        plan: account?.planTier ?? 'free',
        subscriptionStatus: sub?.status ?? null,
        credits: total,
        creditsUsed: usedBy.get(id) ?? 0,
        prompts: use?.prompts ?? 0,
        costUsd: use?.costUsd ?? 0,
        lastPromptAt: use?.last ? new Date(use.last).toISOString() : null,
      };
    });

    if (filters.plan) rows = rows.filter((r) => r.plan === filters.plan);

    const field = SORT_FIELD[filters.sort] as keyof AdminUserRow;
    rows.sort((a, b) => {
      const av = a[field];
      const bv = b[field];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return bv - av;
      return String(bv).localeCompare(String(av));
    });

    const filteredTotal = filters.plan ? rows.length : total;
    const start = (filters.page - 1) * USERS_PAGE_SIZE;
    return { rows: rows.slice(start, start + USERS_PAGE_SIZE), total: filteredTotal };
  }

  async getUser(id: string): Promise<Record<string, unknown> | null> {
    const db = this.connection.db;
    if (!db || !ObjectId.isValid(id) || !/^[0-9a-f]{24}$/i.test(id)) return null;
    const objectId = new ObjectId(id);
    const doc = await db.collection('user').findOne({ _id: objectId });
    if (!doc) return null;

    const since30 = new Date(Date.now() - 30 * 86_400_000);
    const [account, subs, ledger, sessionAgg, conversations, oauth, usedAll, used30] = await Promise.all([
      this.creditAccountModel.findOne({ billingEntityId: id }).lean(),
      this.subscriptionModel.find({ billingEntityId: id }).sort({ createdAt: -1 }).lean(),
      this.creditLedgerModel.find({ billingEntityId: id }).sort({ createdAt: -1 }).limit(25).lean(),
      db
        .collection('session')
        .aggregate([{ $match: { userId: objectId } }, { $group: { _id: null, count: { $sum: 1 }, last: { $max: '$updatedAt' } } }])
        .toArray(),
      db.collection('conversations').countDocuments({ userId: id }),
      db.collection('account').find({ userId: objectId }, { projection: { providerId: 1 } }).toArray(),
      this.sumCreditsUsed(id),
      this.sumCreditsUsed(id, since30),
    ]);

    return {
      user: {
        id,
        name: doc.name ?? '',
        email: doc.email ?? '',
        image: doc.image ?? null,
        emailVerified: Boolean(doc.emailVerified),
        createdAt: doc.createdAt ? new Date(doc.createdAt).toISOString() : null,
        providers: [...new Set(oauth.map((a) => a.providerId as string))],
      },
      sessions: {
        count: num((sessionAgg[0] as { count?: number } | undefined)?.count),
        lastSeenAt: (sessionAgg[0] as { last?: Date } | undefined)?.last
          ? new Date((sessionAgg[0] as { last: Date }).last).toISOString()
          : null,
      },
      conversations,
      balance: account
        ? {
            planTier: account.planTier ?? 'free',
            planCredits: num(account.planCredits),
            purchasedCredits: num(account.purchasedCredits),
            oneTimeCredits: num(account.oneTimeCredits),
            total: num(account.planCredits) + num(account.purchasedCredits) + num(account.oneTimeCredits),
            currentPeriodEnd: account.currentPeriodEnd ? new Date(account.currentPeriodEnd).toISOString() : null,
          }
        : null,
      subscriptions: subs.map((s) => ({
        id: String(s._id),
        planTier: s.planTier,
        status: s.status,
        currentPeriodEnd: s.currentPeriodEnd ? new Date(s.currentPeriodEnd).toISOString() : null,
        cancelAtPeriodEnd: Boolean(s.cancelAtPeriodEnd),
        createdAt: s.createdAt ? new Date(s.createdAt).toISOString() : null,
      })),
      ledger: ledger.map((doc2) => ({
        id: String(doc2._id),
        billingEntityId: doc2.billingEntityId,
        entryType: doc2.entryType,
        amount: num(doc2.amount),
        bucket: doc2.bucket,
        actionType: doc2.actionType ?? null,
        conversationId: doc2.conversationId ?? null,
        createdAt: doc2.createdAt ? new Date(doc2.createdAt).toISOString() : null,
      })),
      recentPrompts: await this.adminPrompts.listRecentByUser(id, RECENT_PROMPTS_LIMIT),
      usage: { allTime: { ...(await this.summarize(id)), creditsUsed: usedAll }, last30: { ...(await this.summarize(id, since30)), creditsUsed: used30 } },
    };
  }

  private async sumCreditsUsed(userId: string, since?: Date): Promise<number> {
    const [row] = await this.creditLedgerModel.aggregate<{ credits: number }>([
      { $match: { billingEntityId: userId, entryType: 'debit', ...(since ? { createdAt: { $gte: since } } : {}) } },
      { $group: { _id: null, credits: { $sum: { $multiply: ['$amount', -1] } } } },
    ]);
    return num(row?.credits);
  }

  private async summarize(userId: string, since?: Date) {
    const db = this.connection.db;
    if (!db) return { prompts: 0, calls: 0, tokens: 0, costUsd: 0 };
    const [row] = await db
      .collection('ai_prompts')
      .aggregate([
        { $match: { userId, ...(since ? { createdAt: { $gte: since } } : {}) } },
        { $group: { _id: null, prompts: { $sum: 1 }, calls: { $sum: '$llmCalls' }, tokens: { $sum: '$totalTokens' }, costUsd: { $sum: '$costUsd' } } },
      ])
      .toArray();
    return { prompts: num(row?.prompts), calls: num(row?.calls), tokens: num(row?.tokens), costUsd: num(row?.costUsd) };
  }
}
