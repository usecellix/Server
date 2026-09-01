import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { TierAMetricsService } from '../src/audit/tier-a-metrics.service';
import { WorkflowTraceDocument } from '../src/common/logging/schemas/workflow-trace.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Same in-memory-Map-backed fake Mongoose model convention as #20/#29's specs — real
 * enough `create`/`find`/`findOneAndUpdate` semantics to exercise the actual
 * ChangeSetService + TierAMetricsService pipeline (query -> join -> compute), not just
 * tier-a-metrics.util.ts's pure functions in isolation (those are covered directly in
 * tier-a-metrics.util.spec.ts).
 */
function createChangeSetModel() {
  const store = new Map<string, Record<string, unknown> & { save: jest.Mock }>();

  function withExec<T>(value: T) {
    const promise = Promise.resolve(value);
    (promise as Promise<T> & { exec: () => Promise<T> }).exec = () => promise;
    return promise;
  }

  return {
    create: jest.fn(async (data: Record<string, unknown>) => {
      const doc = {
        ...data,
        save: jest.fn(async function (this: Record<string, unknown>) {
          store.set(this.changeSetId as string, this as Record<string, unknown> & { save: jest.Mock });
        }),
      } as Record<string, unknown> & { save: jest.Mock };
      store.set(doc.changeSetId as string, doc);
      return doc;
    }),
    findOne: jest.fn(({ changeSetId }: { changeSetId: string }) =>
      withExec(store.get(changeSetId) ?? null),
    ),
    findOneAndUpdate: jest.fn(
      (
        { changeSetId, status }: { changeSetId: string; status: string },
        update: Record<string, unknown>,
      ) => {
        const doc = store.get(changeSetId);
        if (!doc || doc.status !== status) return withExec(null);
        Object.assign(doc, update);
        return withExec(doc);
      },
    ),
    find: jest.fn((query: { timestamp?: { $gte: Date; $lte: Date } }) => {
      const all = [...store.values()].filter((doc) => {
        if (!query.timestamp) return true;
        const ts = doc.timestamp as Date;
        return ts >= query.timestamp.$gte && ts <= query.timestamp.$lte;
      });
      return {
        sort: () => ({
          limit: () => ({ exec: () => Promise.resolve(all) }),
        }),
      };
    }),
  };
}

function createConversationModel() {
  return {
    findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })),
  };
}

interface FakeTraceDoc {
  changeSetId?: string;
  route?: string;
  tier?: number;
  status: string;
  ts: Date;
  nodes: { type: string; status: string }[];
}

function createWorkflowTraceModel(traces: FakeTraceDoc[]) {
  return {
    find: jest.fn((query: { ts?: { $gte: Date; $lte: Date }; changeSetId?: { $in: string[] } }) => {
      let results = traces;
      if (query.ts) {
        results = results.filter((t) => t.ts >= query.ts!.$gte && t.ts <= query.ts!.$lte);
      }
      if (query.changeSetId) {
        const ids = new Set(query.changeSetId.$in);
        results = results.filter((t) => t.changeSetId && ids.has(t.changeSetId));
      }
      return {
        select: () => ({
          lean: () => ({ exec: () => Promise.resolve(results) }),
        }),
      };
    }),
  };
}

function buildChangeSetService() {
  const changeSetModel = createChangeSetModel();
  const conversationModel = createConversationModel();
  const workflowTrace = {
    appendTerminalByChangeSet: jest.fn(),
    appendTerminalByConversationId: jest.fn(),
  } as unknown as WorkflowTraceService;
  return new ChangeSetService(
    changeSetModel as unknown as Model<ChangeSetDocument>,
    conversationModel as never,
    workflowTrace,
  );
}

function buildTierAMetricsService(changeSetService: ChangeSetService, traces: FakeTraceDoc[]) {
  const traceModel = createWorkflowTraceModel(traces);
  return new TierAMetricsService(
    traceModel as unknown as Model<WorkflowTraceDocument>,
    changeSetService,
  );
}

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B2',
      rowCount: 2,
      columnCount: 2,
      values: [
        ['Item', 'Total'],
        ['Apple', 10],
      ],
      formulas: [['', ''], ['', '']],
      numberFormats: [['General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('TierAMetricsService (TASKS.md #50 — real query/join pipeline)', () => {
  // A window around "now" so a real ChangeSetService.createPreview() call's own
  // `timestamp: new Date()` naturally lands inside it, no timestamp hacking needed.
  const from = new Date(Date.now() - 60 * 60 * 1000);
  const to = new Date(Date.now() + 60 * 60 * 1000);
  const inRange = new Date();

  it('joins an applied change set to its trace via changeSetId to resolve route/tier for A5', async () => {
    const changeSetService = buildChangeSetService();

    const preview = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'Bump the total to something unusual',
      context,
      actions: [{ type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: '#REF!' } as Action],
    });
    await changeSetService.markApplied(preview.changeSetId);

    // The trace fixture is only knowable *after* createPreview mints a real changeSetId —
    // this is the actual join key TierAMetricsService uses to resolve route/tier for a
    // change set, which change_sets itself doesn't carry.
    const tierAMetricsService = buildTierAMetricsService(changeSetService, [
      {
        changeSetId: preview.changeSetId,
        route: 'write',
        tier: 2,
        status: 'accepted',
        ts: inRange,
        nodes: [{ type: 'verifier', status: 'success' }],
      },
    ]);

    const report = await tierAMetricsService.getReport(from, to);
    expect(report.segments['write/tier2'].a5UnintendedModificationRate).toEqual({
      rate: 0,
      numerator: 0,
      denominator: 1,
    });
    // The literal error injected above should be attributed to the correct segment too.
    expect(report.segments['write/tier2'].a6FormulaErrorRate).toEqual({
      rate: 1,
      numerator: 1,
      denominator: 1,
    });
  });

  it('computes A1/A4 from workflow_traces alone (no change sets in range)', async () => {
    const changeSetService = buildChangeSetService();
    const tierAMetricsService = buildTierAMetricsService(changeSetService, [
      { route: 'write', tier: 3, status: 'accepted', ts: inRange, nodes: [] }, // no verifier node -> skipped -> false success
      { route: 'write', tier: 3, status: 'accepted', ts: inRange, nodes: [{ type: 'verifier', status: 'success' }] },
    ]);

    const report = await tierAMetricsService.getReport(from, to);
    expect(report.overall.a1FalseSuccessRate).toEqual({ rate: 0.5, numerator: 1, denominator: 2 });
    expect(report.segments['write/tier3'].a1FalseSuccessRate).toEqual({
      rate: 0.5,
      numerator: 1,
      denominator: 2,
    });
  });

  it('excludes traces outside the requested date range', async () => {
    const outOfRange = new Date('2020-01-01T00:00:00Z');
    const changeSetService = buildChangeSetService();
    const tierAMetricsService = buildTierAMetricsService(changeSetService, [
      { route: 'write', tier: 1, status: 'accepted', ts: outOfRange, nodes: [] },
    ]);

    const report = await tierAMetricsService.getReport(from, to);
    expect(report.overall.a1FalseSuccessRate).toEqual({ rate: null, numerator: 0, denominator: 0 });
    expect(Object.keys(report.segments)).toEqual([]);
  });
});
