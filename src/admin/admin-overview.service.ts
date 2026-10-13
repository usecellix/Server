import { Injectable } from '@nestjs/common';
import { InjectConnection } from '@nestjs/mongoose';
import { Connection } from 'mongoose';
import { fillBuckets, num, PLAN_PRICE_INR, TZ, usersById } from './admin-common';

export interface OverviewRange {
  from: Date | null;
  to: Date;
  bucket: 'hour' | 'day';
}

/** Port of Dashboard/src/lib/data/overview.ts's getOverview. */
@Injectable()
export class AdminOverviewService {
  constructor(@InjectConnection() private readonly connection: Connection) {}

  async getOverview(range: OverviewRange) {
    const db = this.connection.db;
    if (!db) return null;
    const prompts = db.collection('ai_prompts');
    const match: Record<string, unknown> = range.from ? { createdAt: { $gte: range.from } } : {};

    const seriesFrom =
      range.from ??
      (await prompts.find({}, { projection: { createdAt: 1 } }).sort({ createdAt: 1 }).limit(1).next())?.createdAt ??
      new Date(range.to.getTime() - 30 * 86_400_000);

    const [totalsAgg, seriesAgg, topUsersAgg, modelAgg, recent, usersTotal, usersNew, activeSubs] = await Promise.all([
      prompts
        .aggregate([
          { $match: match },
          {
            $group: {
              _id: null,
              prompts: { $sum: 1 },
              llmCalls: { $sum: '$llmCalls' },
              failedCalls: { $sum: '$failedCalls' },
              totalTokens: { $sum: '$totalTokens' },
              costUsd: { $sum: '$costUsd' },
              erroredPrompts: { $sum: { $cond: [{ $eq: ['$lastOutcome', 'error'] }, 1, 0] } },
              users: { $addToSet: '$userId' },
            },
          },
        ])
        .toArray(),
      prompts
        .aggregate<{ _id: Date; cost: number; count: number }>([
          { $match: { createdAt: { $gte: seriesFrom } } },
          {
            $group: {
              _id: { $dateTrunc: { date: '$createdAt', unit: range.bucket, timezone: TZ } },
              cost: { $sum: '$costUsd' },
              count: { $sum: 1 },
            },
          },
        ])
        .toArray(),
      prompts
        .aggregate<{ _id: string | null; prompts: number; costUsd: number; tokens: number }>([
          { $match: match },
          { $group: { _id: '$userId', prompts: { $sum: 1 }, costUsd: { $sum: '$costUsd' }, tokens: { $sum: '$totalTokens' } } },
          { $sort: { costUsd: -1 } },
          { $limit: 6 },
        ])
        .toArray(),
      db
        .collection('llm_calls')
        .aggregate<{ _id: string; calls: number; costUsd: number; tokens: number }>([
          { $match: range.from ? { ts: { $gte: range.from } } : {} },
          { $group: { _id: '$model', calls: { $sum: 1 }, costUsd: { $sum: '$costUsd' }, tokens: { $sum: '$totalTokens' } } },
          { $sort: { costUsd: -1 } },
          { $limit: 6 },
        ])
        .toArray(),
      prompts.find(match).sort({ createdAt: -1 }).limit(8).toArray(),
      db.collection('user').countDocuments(),
      db.collection('user').countDocuments(match),
      db
        .collection('subscriptions')
        .aggregate<{ _id: string; count: number }>([{ $match: { status: 'active' } }, { $group: { _id: '$planTier', count: { $sum: 1 } } }])
        .toArray(),
    ]);

    const t = (totalsAgg[0] ?? {}) as Record<string, unknown>;
    const people = await usersById(this.connection, [
      ...topUsersAgg.map((u) => u._id ?? undefined),
      ...recent.map((p) => p.userId as string | undefined),
    ]);
    const recentPromptIds = recent.map((p) => p.promptId as string);
    const charged = await this.creditsChargedByPrompt(recentPromptIds);

    const byPlan: Record<string, number> = {};
    let mrrInr = 0;
    for (const row of activeSubs) {
      byPlan[row._id] = row.count;
      mrrInr += (PLAN_PRICE_INR[row._id] ?? 0) * row.count;
    }

    return {
      totals: {
        prompts: num(t.prompts),
        llmCalls: num(t.llmCalls),
        failedCalls: num(t.failedCalls),
        totalTokens: num(t.totalTokens),
        costUsd: num(t.costUsd),
        erroredPrompts: num(t.erroredPrompts),
        activeUsers: Array.isArray(t.users) ? (t.users as unknown[]).filter(Boolean).length : 0,
      },
      spendSeries: fillBuckets(range.bucket, seriesFrom, range.to, seriesAgg.map((r) => ({ _id: r._id, value: r.cost }))),
      promptSeries: fillBuckets(range.bucket, seriesFrom, range.to, seriesAgg.map((r) => ({ _id: r._id, value: r.count }))),
      topUsers: topUsersAgg.map((row) => ({
        userId: row._id,
        user: row._id ? (people.get(row._id) ?? null) : null,
        prompts: row.prompts,
        costUsd: row.costUsd,
        tokens: row.tokens,
      })),
      topModels: modelAgg.map((row) => ({ model: row._id, calls: row.calls, costUsd: row.costUsd, tokens: row.tokens })),
      users: { total: usersTotal, newInRange: usersNew },
      subscriptions: { active: activeSubs.reduce((sum, r) => sum + r.count, 0), mrrInr, byPlan },
      recentPrompts: recent.map((doc) => this.toPromptRowLite(doc, people, charged.get(doc.promptId as string) ?? 0)),
    };
  }

  /** Same row shape AdminPromptsService.toPromptRow builds — duplicated (not shared) since that method is private there; keep in sync if PromptRow's shape changes. */
  private toPromptRowLite(
    doc: Record<string, unknown>,
    people: Map<string, { id: string; name: string; email: string; image: string | null }>,
    creditsCharged: number,
  ) {
    const userId = (doc.userId as string | undefined) ?? null;
    return {
      promptId: doc.promptId,
      prompt: doc.prompt ?? '',
      userId,
      user: userId ? (people.get(userId) ?? null) : null,
      conversationId: doc.conversationId ?? null,
      mode: doc.mode ?? null,
      route: doc.route ?? null,
      tier: typeof doc.tier === 'number' ? doc.tier : null,
      createdAt: new Date(doc.createdAt as string | Date).toISOString(),
      lastActivityAt: new Date((doc.lastActivityAt ?? doc.createdAt) as string | Date).toISOString(),
      requestCount: num(doc.requestCount),
      llmCalls: num(doc.llmCalls),
      failedCalls: num(doc.failedCalls),
      promptTokens: num(doc.promptTokens),
      completionTokens: num(doc.completionTokens),
      reasoningTokens: num(doc.reasoningTokens),
      cachedTokens: num(doc.cachedTokens),
      totalTokens: num(doc.totalTokens),
      costUsd: num(doc.costUsd),
      creditsCharged,
      llmLatencyMs: num(doc.llmLatencyMs),
      requestDurationMs: num(doc.requestDurationMs),
      models: Array.isArray(doc.models) ? doc.models : [],
      outcome: doc.lastOutcome === 'error' ? 'error' : doc.lastOutcome === 'ok' ? 'ok' : 'running',
      lastError: doc.lastError ?? null,
    };
  }

  private async creditsChargedByPrompt(promptIds: string[]): Promise<Map<string, number>> {
    const db = this.connection.db;
    const ids = [...new Set(promptIds.filter(Boolean))];
    if (!db || ids.length === 0) return new Map();
    const rows = await db
      .collection('credit_ledger')
      .aggregate<{ _id: string; credits: number }>([
        { $match: { promptId: { $in: ids }, entryType: 'debit' } },
        { $group: { _id: '$promptId', credits: { $sum: { $multiply: ['$amount', -1] } } } },
      ])
      .toArray();
    return new Map(rows.map((r) => [r._id, num(r.credits)]));
  }
}
