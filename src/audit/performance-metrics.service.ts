import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import {
  WorkflowTrace,
  WorkflowTraceDocument,
} from '../common/logging/schemas/workflow-trace.schema';
import { AuditLog, AuditLogDocument } from './schemas/audit-log.schema';
import { ContextCacheService } from '../common/cache/context-cache.service';
import {
  buildPerformanceReport,
  PerformanceReport,
  TokenUsageRecord,
  TraceLatencyRecord,
} from './performance-metrics.util';

@Injectable()
export class PerformanceMetricsService {
  constructor(
    // Read-only — same cross-module registration pattern as TierAMetricsService.
    @InjectModel(WorkflowTrace.name)
    private readonly workflowTraceModel: Model<WorkflowTraceDocument>,
    @InjectModel(AuditLog.name)
    private readonly auditLogModel: Model<AuditLogDocument>,
    private readonly contextCache: ContextCacheService,
  ) {}

  async getReport(fromDate: Date, toDate: Date): Promise<PerformanceReport> {
    const [traces, auditLogs] = await Promise.all([
      this.workflowTraceModel
        .find({ ts: { $gte: fromDate, $lte: toDate } })
        .select({ route: 1, tier: 1, durationMs: 1, nodes: 1 })
        .lean()
        .exec(),
      this.auditLogModel
        .find({ timestamp: { $gte: fromDate, $lte: toDate } })
        .select({ tier: 1, promptTokens: 1, completionTokens: 1 })
        .lean()
        .exec(),
    ]);

    const traceRecords: TraceLatencyRecord[] = traces.map((trace) => ({
      route: trace.route,
      tier: trace.tier,
      // Prefer the trace-level durationMs set by finalizeWorkflow() at most
      // (13/19, confirmed by direct grep) terminal call sites — it's the
      // request's own measured wall-clock time, no derivation needed. Fall
      // back to the node-span derivation only for the remaining call sites
      // that don't set it (e.g. some clarification/error paths), rather than
      // dropping those traces from latency reporting entirely.
      totalDurationMs: trace.durationMs ?? computeTraceDurationMs(trace.nodes ?? []),
    }));

    const tokenRecords: TokenUsageRecord[] = auditLogs.map((log) => ({
      tier: log.tier,
      promptTokens: log.promptTokens,
      completionTokens: log.completionTokens,
    }));

    const cacheStats = this.contextCache.getStats();

    return buildPerformanceReport(traceRecords, tokenRecords, {
      hits: cacheStats.hits,
      misses: cacheStats.misses,
      stableHits: cacheStats.stableHits,
      stableMisses: cacheStats.stableMisses,
    });
  }
}

/**
 * Derives one trace's total wall-clock span from its nodes' timestamps —
 * min(startedAt) to max(endedAt) across all nodes that have both set. Returns
 * null (not 0) when no node has a usable timestamp pair, so a trace with
 * incomplete instrumentation is excluded from percentile computation rather
 * than silently counted as an instant (0ms) request — see
 * performance-metrics.util.ts's module doc for why span, not summed
 * durationMs, is the correct aggregation.
 */
function computeTraceDurationMs(
  nodes: { startedAt?: Date; endedAt?: Date }[],
): number | null {
  let minStart: number | null = null;
  let maxEnd: number | null = null;

  for (const node of nodes) {
    if (node.startedAt) {
      const start = new Date(node.startedAt).getTime();
      if (minStart === null || start < minStart) minStart = start;
    }
    if (node.endedAt) {
      const end = new Date(node.endedAt).getTime();
      if (maxEnd === null || end > maxEnd) maxEnd = end;
    }
  }

  if (minStart === null || maxEnd === null || maxEnd < minStart) return null;
  return maxEnd - minStart;
}
