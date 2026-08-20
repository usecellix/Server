import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import {
  WorkflowNode,
  WorkflowTrace,
  WorkflowTraceDocument,
} from '../common/logging/schemas/workflow-trace.schema';
import { ChangeSetService } from './change-set.service';
import {
  buildTierAMetricsReport,
  ChangeSetOutcomeRecord,
  TierAMetricsReport,
  TraceOutcomeRecord,
} from './tier-a-metrics.util';

@Injectable()
export class TierAMetricsService {
  constructor(
    // Read-only — WorkflowTrace is owned/written by LoggingModule (TASKS.md #24's same
    // cross-module registration pattern this module already uses for `Conversation`).
    @InjectModel(WorkflowTrace.name)
    private readonly workflowTraceModel: Model<WorkflowTraceDocument>,
    private readonly changeSetService: ChangeSetService,
  ) {}

  async getReport(fromDate: Date, toDate: Date): Promise<TierAMetricsReport> {
    const traces = await this.workflowTraceModel
      .find({ ts: { $gte: fromDate, $lte: toDate } })
      .select({ route: 1, tier: 1, status: 1, nodes: 1, changeSetId: 1 })
      .lean()
      .exec();

    const traceRecords: TraceOutcomeRecord[] = traces.map((trace) => ({
      route: trace.route,
      tier: trace.tier,
      status: trace.status,
      verifierNodeStatus: resolveVerifierNodeStatus(trace.nodes ?? []),
    }));

    // A5/A6 only count *applied* change sets (PRD's own wording: "applied change sets
    // containing..."), matching change-set.service.ts's own durable, no-TTL query.
    const changeSets = (await this.changeSetService.getByDateRange(fromDate, toDate)).filter(
      (cs) => cs.status === 'applied',
    );

    // change_sets doesn't carry route/tier itself — only its originating workflow_trace
    // does. A second, targeted query by changeSetId (not by the same ts window, since a
    // change set's own timestamp and its trace's ts can drift slightly) rather than
    // reusing `traces` above, so this join is correct even when the trace fell just
    // outside the requested date range.
    const changeSetIds = changeSets.map((cs) => cs.changeSetId);
    const routeTierByChangeSetId = new Map<string, { route?: string; tier?: number }>();
    if (changeSetIds.length > 0) {
      const linkedTraces = await this.workflowTraceModel
        .find({ changeSetId: { $in: changeSetIds } })
        .select({ changeSetId: 1, route: 1, tier: 1 })
        .lean()
        .exec();
      for (const trace of linkedTraces) {
        if (trace.changeSetId) {
          routeTierByChangeSetId.set(trace.changeSetId, { route: trace.route, tier: trace.tier });
        }
      }
    }

    const changeSetRecords: ChangeSetOutcomeRecord[] = changeSets.map((cs) => {
      const linked = routeTierByChangeSetId.get(cs.changeSetId);
      return {
        route: linked?.route,
        tier: linked?.tier,
        unintendedCount: cs.unintendedChanges.length,
        formulaErrorCount: cs.formulaErrorsIntroduced.length,
      };
    });

    return buildTierAMetricsReport(traceRecords, changeSetRecords);
  }
}

function resolveVerifierNodeStatus(nodes: WorkflowNode[]): 'success' | 'failed' | 'skipped' {
  const verifierNode = nodes.find((node) => node.type === 'verifier');
  if (!verifierNode) return 'skipped';
  return verifierNode.status === 'success' ? 'success' : 'failed';
}
