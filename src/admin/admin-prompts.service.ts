import { Injectable } from '@nestjs/common';
import { InjectConnection } from '@nestjs/mongoose';
import { Connection } from 'mongoose';
import { AppConfigService } from '../config/app-config.service';
import { escapeRegex, num, usersById } from './admin-common';

const PROMPTS_PAGE_SIZE = 25;

export type PromptSort = 'recent' | 'cost' | 'tokens' | 'calls';
const PROMPT_SORT: Record<PromptSort, Record<string, 1 | -1>> = {
  recent: { createdAt: -1 },
  cost: { costUsd: -1 },
  tokens: { totalTokens: -1 },
  calls: { llmCalls: -1 },
};

/**
 * Port of `Dashboard/src/lib/data/prompts.ts` — same shapes, same queries,
 * moved onto the backend's own Mongo connection (AdminGuard's docblock
 * explains why). Reads `ai_prompts`, `llm_calls`, `credit_ledger`, and
 * `user` directly (no Mongoose schemas for these; ai_prompts/llm_calls
 * predate this module and credit_ledger already has one, but going through
 * the raw connection keeps this service's queries a 1:1 mirror of the
 * Dashboard's originals rather than reshaping them around the schema).
 */
@Injectable()
export class AdminPromptsService {
  constructor(
    @InjectConnection() private readonly connection: Connection,
    private readonly config: AppConfigService,
  ) {}

  /**
   * The `recentPrompts` list on a user's detail page — the original
   * `Dashboard/src/lib/data/users.ts` `getUser` queried `ai_prompts` for this
   * directly (`.limit(15)`, no sums/credit-ledger $lookup) rather than
   * reusing `listPrompts`'s heavier page+sums query. Kept as its own method
   * here for the same reason: this doesn't need a total or a range sum.
   */
  async listRecentByUser(userId: string, limit: number) {
    const db = this.connection.db;
    if (!db) return [];
    const docs = await db.collection('ai_prompts').find({ userId }).sort({ createdAt: -1 }).limit(limit).toArray();
    if (docs.length === 0) return [];
    const people = await usersById(this.connection, [userId]);
    const charged = await this.creditsChargedByPrompt(docs.map((d) => d.promptId as string));
    return docs.map((d) => this.toPromptRow(d, people, charged.get(d.promptId as string) ?? 0));
  }

  async listPrompts(filters: {
    since?: Date;
    userId?: string;
    status?: 'error' | 'ok' | 'running';
    q?: string;
    sort: PromptSort;
    page: number;
  }) {
    const db = this.connection.db;
    if (!db) return { rows: [], total: 0, sums: { costUsd: 0, tokens: 0, calls: 0, credits: 0 } };

    const where: Record<string, unknown> = {};
    if (filters.since) where.createdAt = { $gte: filters.since };
    if (filters.userId) where.userId = filters.userId;
    if (filters.status) where.lastOutcome = filters.status;
    if (filters.q) where.prompt = { $regex: escapeRegex(filters.q.slice(0, 200)), $options: 'i' };

    const prompts = db.collection('ai_prompts');
    const [docs, total, sums] = await Promise.all([
      prompts
        .find(where)
        .sort({ ...PROMPT_SORT[filters.sort], _id: -1 })
        .skip((filters.page - 1) * PROMPTS_PAGE_SIZE)
        .limit(PROMPTS_PAGE_SIZE)
        .toArray(),
      prompts.countDocuments(where),
      prompts
        .aggregate([
          { $match: where },
          { $lookup: { from: 'credit_ledger', localField: 'promptId', foreignField: 'promptId', as: 'ledger' } },
          {
            $group: {
              _id: null,
              costUsd: { $sum: '$costUsd' },
              tokens: { $sum: '$totalTokens' },
              calls: { $sum: '$llmCalls' },
              credits: {
                $sum: {
                  $sum: {
                    $map: {
                      input: { $filter: { input: '$ledger', cond: { $eq: ['$$this.entryType', 'debit'] } } },
                      in: { $multiply: ['$$this.amount', -1] },
                    },
                  },
                },
              },
            },
          },
        ])
        .toArray(),
    ]);

    const people = await usersById(this.connection, docs.map((d) => d.userId as string | undefined));
    const charged = await this.creditsChargedByPrompt(docs.map((d) => d.promptId as string));
    const s = sums[0] ?? ({} as Record<string, number>);
    return {
      rows: docs.map((d) => this.toPromptRow(d, people, charged.get(d.promptId as string) ?? 0)),
      total,
      sums: { costUsd: num(s.costUsd), tokens: num(s.tokens), calls: num(s.calls), credits: num(s.credits) },
    };
  }

  async getPrompt(promptId: string) {
    const db = this.connection.db;
    if (!db) return null;
    const doc = await db.collection('ai_prompts').findOne({ promptId });
    if (!doc) return null;

    const callDocs = await db.collection('llm_calls').find({ promptId }).sort({ ts: 1 }).limit(2000).toArray();
    const people = await usersById(this.connection, [doc.userId as string | undefined]);
    const creditsPerUsd = this.config.creditsPerUsd;
    const calls = callDocs.map((c) => {
      const ts = new Date(c.ts);
      const latencyMs = num(c.latencyMs);
      return {
        id: String(c._id),
        ts: ts.toISOString(),
        startedAt: new Date(ts.getTime() - latencyMs).toISOString(),
        model: c.model,
        servedModel: c.servedModel ?? null,
        caller: c.caller ?? 'unknown',
        attempt: num(c.attempt) || 1,
        streaming: Boolean(c.streaming),
        promptTokens: num(c.promptTokens),
        completionTokens: num(c.completionTokens),
        reasoningTokens: num(c.reasoningTokens),
        cachedTokens: num(c.cachedTokens),
        totalTokens: num(c.totalTokens),
        costUsd: num(c.costUsd),
        creditsEquivalent: num(c.costUsd) * creditsPerUsd,
        costEstimated: Boolean(c.costEstimated),
        latencyMs,
        success: Boolean(c.success),
        finishReason: c.finishReason ?? null,
        errorStatus: typeof c.errorStatus === 'number' ? c.errorStatus : null,
        errorMessage: c.errorMessage ?? null,
      };
    });
    calls.sort((a, b) => a.startedAt.localeCompare(b.startedAt));

    const charged = await this.creditsChargedByPrompt([promptId]);
    return {
      prompt: this.toPromptRow(doc, people, charged.get(promptId) ?? 0),
      calls,
      byCaller: this.breakdown(calls, (c) => c.caller),
      byModel: this.breakdown(calls, (c) => c.servedModel ?? c.model),
      retries: calls.filter((c) => c.attempt > 1).length,
      estimatedCostCalls: calls.filter((c) => c.costEstimated).length,
    };
  }

  private toPromptRow(
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

  private breakdown<T extends { caller: string; costUsd: number; totalTokens: number; latencyMs: number; success: boolean }>(
    calls: T[],
    keyOf: (c: T) => string,
  ) {
    const map = new Map<string, { key: string; calls: number; failed: number; tokens: number; costUsd: number; latencyMs: number }>();
    for (const call of calls) {
      const key = keyOf(call);
      const row = map.get(key) ?? { key, calls: 0, failed: 0, tokens: 0, costUsd: 0, latencyMs: 0 };
      row.calls += 1;
      row.failed += call.success ? 0 : 1;
      row.tokens += call.totalTokens;
      row.costUsd += call.costUsd;
      row.latencyMs += call.latencyMs;
      map.set(key, row);
    }
    return [...map.values()].sort((a, b) => b.costUsd - a.costUsd || b.calls - a.calls);
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
