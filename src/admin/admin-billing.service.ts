import { Injectable } from '@nestjs/common';
import { InjectConnection } from '@nestjs/mongoose';
import { Connection } from 'mongoose';
import { escapeRegex, num, PLAN_PRICE_INR, resolveOwners } from './admin-common';

const SUBS_PAGE_SIZE = 25;
const LEDGER_PAGE_SIZE = 25;

export type SubSort = 'recent' | 'price' | 'period';
export type LedgerSort = 'recent' | 'amount';

function toLedgerRow(doc: Record<string, unknown>) {
  return {
    id: String(doc._id),
    billingEntityId: doc.billingEntityId,
    entryType: doc.entryType,
    amount: num(doc.amount),
    bucket: doc.bucket,
    actionType: doc.actionType ?? null,
    conversationId: doc.conversationId ?? null,
    createdAt: doc.createdAt ? new Date(doc.createdAt as string | Date).toISOString() : null,
  };
}

/** Port of Dashboard/src/lib/data/billing.ts. */
@Injectable()
export class AdminBillingService {
  constructor(@InjectConnection() private readonly connection: Connection) {}

  async getBillingSummary(since: Date | null) {
    const db = this.connection.db;
    if (!db) return null;
    const match = since ? { createdAt: { $gte: since } } : {};

    const [subs, accounts, ledgerSums] = await Promise.all([
      db.collection('subscriptions').find({}).toArray(),
      db.collection('credit_accounts').find({}).toArray(),
      db
        .collection('credit_ledger')
        .aggregate<{ _id: string; amount: number; count: number }>([
          { $match: match },
          { $group: { _id: '$entryType', amount: { $sum: '$amount' }, count: { $sum: 1 } } },
        ])
        .toArray(),
    ]);

    const owners = await resolveOwners(this.connection, [
      ...subs.map((s) => s.billingEntityId as string),
      ...accounts.map((a) => a.billingEntityId as string),
    ]);

    const statusCounts: Record<string, number> = {};
    const activeByPlan: Record<string, number> = {};
    let mrrInr = 0;
    let activeCount = 0;
    let cancelling = 0;
    for (const s of subs) {
      statusCounts[s.status] = (statusCounts[s.status] ?? 0) + 1;
      if (s.status === 'active') {
        activeCount += 1;
        const price = PLAN_PRICE_INR[s.planTier] ?? 0;
        activeByPlan[s.planTier] = (activeByPlan[s.planTier] ?? 0) + 1;
        mrrInr += price;
        if (s.cancelAtPeriodEnd) cancelling += 1;
      }
    }

    const planDistribution: Record<string, number> = {};
    const credits = { plan: 0, purchased: 0, oneTime: 0 };
    for (const a of accounts) {
      planDistribution[a.planTier] = (planDistribution[a.planTier] ?? 0) + 1;
      credits.plan += num(a.planCredits);
      credits.purchased += num(a.purchasedCredits);
      credits.oneTime += num(a.oneTimeCredits);
    }

    const flows: Record<string, { amount: number; count: number }> = {};
    for (const row of ledgerSums) flows[row._id] = { amount: row.amount, count: row.count };

    return {
      statusCounts,
      activeByPlan,
      mrrInr,
      activeCount,
      cancelling,
      accounts: {
        total: accounts.length,
        guest: accounts.filter((a) => owners.get(a.billingEntityId)?.guest).length,
        planDistribution,
        credits,
      },
      flows,
    };
  }

  async listSubscriptions(filters: { since: Date | null; q?: string; status?: string; plan?: string; sort: SubSort; page: number }) {
    const db = this.connection.db;
    if (!db) return { rows: [], total: 0 };
    const where: Record<string, unknown> = filters.since ? { createdAt: { $gte: filters.since } } : {};
    if (filters.status) where.status = filters.status;
    if (filters.plan) where.planTier = filters.plan;

    const docs = await db.collection('subscriptions').find(where).sort({ createdAt: -1 }).toArray();
    const owners = await resolveOwners(this.connection, docs.map((s) => s.billingEntityId as string));

    let rows = docs.map((s) => ({
      id: String(s._id),
      billingEntityId: s.billingEntityId,
      owner: owners.get(s.billingEntityId)!,
      planTier: s.planTier,
      status: s.status,
      priceInr: PLAN_PRICE_INR[s.planTier] ?? 0,
      currentPeriodEnd: s.currentPeriodEnd ? new Date(s.currentPeriodEnd).toISOString() : null,
      cancelAtPeriodEnd: Boolean(s.cancelAtPeriodEnd),
      createdAt: s.createdAt ? new Date(s.createdAt).toISOString() : null,
    }));

    if (filters.q) {
      const q = filters.q.toLowerCase();
      rows = rows.filter(
        (r) => r.owner.label.toLowerCase().includes(q) || (r.owner.email?.toLowerCase().includes(q) ?? false) || r.billingEntityId.toLowerCase().includes(q),
      );
    }

    const sortField: Record<SubSort, keyof (typeof rows)[number]> = { recent: 'createdAt', price: 'priceInr', period: 'currentPeriodEnd' };
    const sortDir: Record<SubSort, 1 | -1> = { recent: -1, price: -1, period: 1 };
    const field = sortField[filters.sort];
    const dir = sortDir[filters.sort];
    rows.sort((a, b) => {
      const av = a[field] as string | number | null;
      const bv = b[field] as string | number | null;
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });

    const total = rows.length;
    const start = (filters.page - 1) * SUBS_PAGE_SIZE;
    return { rows: rows.slice(start, start + SUBS_PAGE_SIZE), total };
  }

  async listLedger(filters: { since: Date | null; q?: string; entryType?: string; sort: LedgerSort; page: number }) {
    const db = this.connection.db;
    if (!db) return { rows: [], total: 0 };
    const where: Record<string, unknown> = filters.since ? { createdAt: { $gte: filters.since } } : {};
    if (filters.entryType) where.entryType = filters.entryType;
    if (filters.q) {
      const pattern = { $regex: escapeRegex(filters.q.slice(0, 200)), $options: 'i' };
      where.$or = [{ billingEntityId: pattern }, { actionType: pattern }, { conversationId: pattern }];
    }

    const sortField = filters.sort === 'amount' ? 'amount' : 'createdAt';
    const [docs, total] = await Promise.all([
      db
        .collection('credit_ledger')
        .find(where)
        .sort({ [sortField]: -1, _id: -1 })
        .skip((filters.page - 1) * LEDGER_PAGE_SIZE)
        .limit(LEDGER_PAGE_SIZE)
        .toArray(),
      db.collection('credit_ledger').countDocuments(where),
    ]);

    const owners = await resolveOwners(this.connection, docs.map((l) => l.billingEntityId as string));
    const rows = docs.map((doc) => ({ ...toLedgerRow(doc), owner: owners.get(doc.billingEntityId as string)! }));
    return { rows, total };
  }
}
