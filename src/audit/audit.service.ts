import { Injectable, Logger } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { AuditLogEntry, LLMTier, MODEL_CONFIGS } from '../types/cellix.types';
import { ChangeSetService } from './change-set.service';
import { AuditEntry, AuditEntryDocument } from './schemas/audit-entry.schema';
import { AuditLog, AuditLogDocument } from './schemas/audit-log.schema';

export interface CreateAuditEntryInput {
  requestId: string;
  processName: string;
  action: string;
  userId?: string;
  confidence?: number;
  payload?: unknown;
  result?: unknown;
}

export interface LogLLMCallParams {
  traceId: string;
  model: string;
  tier: LLMTier;
  intent: string;
  promptTokens: number;
  completionTokens: number;
  latencyMs: number;
  success: boolean;
  errorCode?: string;
  actionsCount?: number;
  rawUsage?: Record<string, unknown>;
}

@Injectable()
export class AuditService {
  private readonly logger = new Logger(AuditService.name);

  constructor(
    @InjectModel(AuditEntry.name)
    private readonly auditEntryModel: Model<AuditEntryDocument>,
    @InjectModel(AuditLog.name)
    private readonly auditLogModel: Model<AuditLogDocument>,
    private readonly changeSetService: ChangeSetService,
  ) {}

  async create(input: CreateAuditEntryInput): Promise<AuditEntry> {
    return this.auditEntryModel.create(input);
  }

  async findByRequestId(requestId: string): Promise<AuditEntry[]> {
    return this.auditEntryModel
      .find({ requestId })
      .sort({ createdAt: 1 })
      .lean<AuditEntry[]>()
      .exec();
  }

  async logLLMCall(params: LogLLMCallParams): Promise<void> {
    const {
      traceId,
      model,
      tier,
      intent,
      promptTokens,
      completionTokens,
      latencyMs,
      success,
      errorCode,
      actionsCount,
      rawUsage,
    } = params;

    const totalTokens = promptTokens + completionTokens;
    const config = MODEL_CONFIGS[tier];

    const estimatedCostUsd =
      (promptTokens / 1000) * config.costPer1kPrompt +
      (completionTokens / 1000) * config.costPer1kCompletion;

    try {
      await this.auditLogModel.create({
        traceId,
        llmModel: model,
        tier,
        intent,
        promptTokens,
        completionTokens,
        totalTokens,
        estimatedCostUsd,
        latencyMs,
        success,
        errorCode,
        actionsCount,
        rawUsage,
      });

      this.logger.log({
        event: 'LLM_CALL',
        traceId,
        model,
        tier,
        intent,
        promptTokens,
        completionTokens,
        totalTokens,
        estimatedCostUsd: estimatedCostUsd.toFixed(6),
        latencyMs,
        success,
        actionsCount,
        ...(errorCode ? { errorCode } : {}),
      });
    } catch (err) {
      this.logger.error('Failed to write audit log', { traceId, err });
    }
  }

  async getLogs(
    options: {
      limit?: number;
      offset?: number;
      tier?: LLMTier;
      fromDate?: Date;
      toDate?: Date;
    } = {},
  ): Promise<{ logs: AuditLogEntry[]; total: number }> {
    const { limit = 50, offset = 0, tier, fromDate, toDate } = options;

    const where: {
      tier?: string;
      timestamp?: { $gte?: Date; $lte?: Date };
    } = {};
    if (tier) where.tier = tier;
    if (fromDate || toDate) {
      where.timestamp = {};
      if (fromDate) where.timestamp.$gte = fromDate;
      if (toDate) where.timestamp.$lte = toDate;
    }

    const [logs, total] = await Promise.all([
      this.auditLogModel
        .find(where)
        .sort({ timestamp: -1 })
        .skip(offset)
        .limit(limit)
        .select({
          traceId: 1,
          timestamp: 1,
          llmModel: 1,
          tier: 1,
          intent: 1,
          promptTokens: 1,
          completionTokens: 1,
          totalTokens: 1,
          estimatedCostUsd: 1,
          latencyMs: 1,
          success: 1,
          errorCode: 1,
          actionsCount: 1,
        })
        .lean()
        .exec(),
      this.auditLogModel.countDocuments(where).exec(),
    ]);

    return {
      logs: logs.map((entry) => ({
        id: String(entry._id),
        traceId: entry.traceId,
        timestamp: entry.timestamp.toISOString(),
        model: entry.llmModel,
        tier: entry.tier as LLMTier,
        intent: entry.intent,
        promptTokens: entry.promptTokens,
        completionTokens: entry.completionTokens,
        totalTokens: entry.totalTokens,
        estimatedCostUsd: entry.estimatedCostUsd,
        latencyMs: entry.latencyMs,
        success: entry.success,
        errorCode: entry.errorCode ?? undefined,
        actionsCount: entry.actionsCount ?? undefined,
      })),
      total,
    };
  }

  async getStats(
    fromDate: Date,
    toDate: Date,
  ): Promise<{
    totalCost: number;
    totalTokens: number;
    totalCalls: number;
    successRate: number;
    avgLatencyMs: number;
    byTier: Record<LLMTier, { calls: number; cost: number; tokens: number }>;
  }> {
    const logs = await this.auditLogModel
      .find({ timestamp: { $gte: fromDate, $lte: toDate } })
      .select({
        tier: 1,
        success: 1,
        totalTokens: 1,
        estimatedCostUsd: 1,
        latencyMs: 1,
      })
      .lean()
      .exec();

    const totalCost = logs.reduce((sum, entry) => sum + entry.estimatedCostUsd, 0);
    const totalTokens = logs.reduce((sum, entry) => sum + entry.totalTokens, 0);
    const totalCalls = logs.length;
    const successCount = logs.filter((entry) => entry.success).length;
    const successRate = totalCalls ? successCount / totalCalls : 0;
    const avgLatencyMs = totalCalls
      ? logs.reduce((sum, entry) => sum + entry.latencyMs, 0) / totalCalls
      : 0;

    const byTier: Record<LLMTier, { calls: number; cost: number; tokens: number }> = {
      low: { calls: 0, cost: 0, tokens: 0 },
      medium: { calls: 0, cost: 0, tokens: 0 },
      high: { calls: 0, cost: 0, tokens: 0 },
    };

    for (const entry of logs) {
      const tier = entry.tier as LLMTier;
      if (!byTier[tier]) continue;
      byTier[tier].calls += 1;
      byTier[tier].cost += entry.estimatedCostUsd;
      byTier[tier].tokens += entry.totalTokens;
    }

    return { totalCost, totalTokens, totalCalls, successRate, avgLatencyMs, byTier };
  }

  async exportAudit(
    fromDate: Date,
    toDate: Date,
  ): Promise<{
    exportedAt: string;
    from: string;
    to: string;
    changeSets: Awaited<ReturnType<ChangeSetService['getByDateRange']>>;
    auditLogs: AuditLogEntry[];
  }> {
    const [changeSets, logsResult] = await Promise.all([
      this.changeSetService.getByDateRange(fromDate, toDate),
      this.getLogs({ fromDate, toDate, limit: 5000, offset: 0 }),
    ]);

    return {
      exportedAt: new Date().toISOString(),
      from: fromDate.toISOString(),
      to: toDate.toISOString(),
      changeSets,
      auditLogs: logsResult.logs,
    };
  }

  buildExportCsv(
    fromDate: Date,
    toDate: Date,
    payload: Awaited<ReturnType<AuditService['exportAudit']>>,
  ): string {
    const lines = [
      `# Cellix audit export`,
      `# from=${fromDate.toISOString()}`,
      `# to=${toDate.toISOString()}`,
      `# exportedAt=${payload.exportedAt}`,
      '',
      'recordType,id,traceId,timestamp,details,costUsd,tokens,status',
    ];

    for (const changeSet of payload.changeSets) {
      const citationSummary = changeSet.changes
        .flatMap((c) => c.sourceRefs ?? [])
        .map((r) => `${r.documentType}:${r.rowOrLine}`)
        .slice(0, 5)
        .join('; ');
      const exceptionSummary = changeSet.changes
        .flatMap((c) => c.exceptionFlags ?? [])
        .map((e) => `${e.severity}:${e.code}`)
        .slice(0, 5)
        .join('; ');
      const details = [
        changeSet.prompt,
        citationSummary ? `sourceRefs=${citationSummary}` : '',
        exceptionSummary ? `exceptionFlags=${exceptionSummary}` : '',
      ]
        .filter(Boolean)
        .join(' | ')
        .replace(/"/g, '""');

      lines.push(
        [
          'change_set',
          changeSet.changeSetId,
          changeSet.traceId,
          new Date(changeSet.timestamp).toISOString(),
          `"${details}"`,
          '',
          '',
          changeSet.status,
        ].join(','),
      );
    }

    for (const log of payload.auditLogs) {
      lines.push(
        [
          'llm_call',
          log.id,
          log.traceId,
          log.timestamp,
          `"${log.intent.replace(/"/g, '""')} (${log.model}/${log.tier})"`,
          log.estimatedCostUsd.toFixed(6),
          log.totalTokens,
          log.success ? 'success' : 'failed',
        ].join(','),
      );
    }

    return lines.join('\n');
  }
}
