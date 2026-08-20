import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { diffShadowsFully } from '../src/audit/diff.engine';

/**
 * Same in-memory fake Mongoose model as change-set.service.conditional-format-revert.spec.ts
 * (TASKS.md #40's precedent, itself reused from #20's e2e test) — real
 * create/findOne/findOneAndUpdate semantics backed by a Map, no live DB.
 */
function createInMemoryModel() {
  const store = new Map<string, Record<string, unknown>>();

  function withExec<T>(value: T) {
    const promise = Promise.resolve(value);
    (promise as Promise<T> & { exec: () => Promise<T> }).exec = () => promise;
    return promise;
  }

  return {
    create: jest.fn(async (data: Record<string, unknown>) => {
      const doc: Record<string, unknown> & { save: jest.Mock } = {
        ...data,
        save: jest.fn(async function (this: Record<string, unknown>) {
          store.set(this.changeSetId as string, this);
        }),
      };
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
  };
}

function createInMemoryConversationModel() {
  return {
    findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })),
  };
}

function buildService() {
  const model = createInMemoryModel();
  const conversationModel = createInMemoryConversationModel();
  const workflowTrace = {
    appendTerminalByChangeSet: jest.fn(),
    appendTerminalByConversationId: jest.fn(),
  } as unknown as WorkflowTraceService;
  return new ChangeSetService(
    model as unknown as Model<ChangeSetDocument>,
    conversationModel as never,
    workflowTrace,
  );
}

const scoresContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B3',
      rowCount: 3,
      columnCount: 2,
      values: [['Item', 'Score'], ['A', 10], ['B', 20]],
      formulas: [['', ''], ['', ''], ['', '']],
      numberFormats: [['General', 'General'], ['General', 'General'], ['General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('ChangeSetService — CREATE_CHART create → apply → revert (TASKS.md #15)', () => {
  it('patches the real chartId in at markApplied time and builds a working DELETE_CHART revert', async () => {
    const action: Action = {
      type: 'CREATE_CHART',
      sheetName: 'Sheet1',
      sourceSheetName: 'Sheet1',
      sourceRange: 'A1:B3',
      chartType: 'column',
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-chart',
      traceId: 'trace-chart',
      prompt: 'chart the scores',
      context: scoresContext,
      actions: [action],
    });

    // Preview time: the structural op exists but has no chartId yet.
    expect(preview.structuralOps).toEqual([
      expect.objectContaining({ opType: 'CREATE_CHART', sheetName: 'Sheet1', params: { sourceRange: 'A1:B3' } }),
    ]);
    expect(preview.irreversibleActionTypes).toEqual([]);

    // Apply-time: the frontend reports back the real Office.js-assigned chart name.
    await service.markApplied(preview.changeSetId, undefined, [
      { sheetName: 'Sheet1', sourceRange: 'A1:B3', chartId: 'Chart-real-42' },
    ]);

    const { inverseActions } = await service.revert(preview.changeSetId);
    expect(inverseActions).toContainEqual({
      type: 'DELETE_CHART',
      sheetName: 'Sheet1',
      chartId: 'Chart-real-42',
    });

    // Independent convergence check, same technique as the #20/#40 e2e tests.
    const original = buildShadowWorkbook(scoresContext);
    const afterForward = virtualApply(original, [action]);
    const afterRevert = virtualApply(afterForward, inverseActions);
    expect(diffShadowsFully(original, afterRevert)).toEqual([]);
  });

  it('leaves the change set unrevertable if markApplied is never given the created id (fails closed)', async () => {
    const action: Action = {
      type: 'CREATE_CHART',
      sheetName: 'Sheet1',
      sourceSheetName: 'Sheet1',
      sourceRange: 'A1:B3',
      chartType: 'column',
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-chart-2',
      traceId: 'trace-chart-2',
      prompt: 'chart the scores',
      context: scoresContext,
      actions: [action],
    });

    await service.markApplied(preview.changeSetId); // no createdChartIds
    await expect(service.revert(preview.changeSetId)).rejects.toThrow(/no chartId was ever reported/);
  });

  it('only patches a structuralOp whose sheetName+sourceRange matches the reported id, in a batch with two creates', async () => {
    const actions: Action[] = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Sheet1',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:B3',
        chartType: 'column',
      } as Action,
      {
        type: 'CREATE_CHART',
        sheetName: 'Sheet1',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:A3',
        chartType: 'pie',
      } as Action,
    ];

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-chart-3',
      traceId: 'trace-chart-3',
      prompt: 'two charts',
      context: scoresContext,
      actions,
    });
    expect(preview.structuralOps).toHaveLength(2);

    // Only report the id for the SECOND chart (A1:A3) this time.
    await service.markApplied(preview.changeSetId, undefined, [
      { sheetName: 'Sheet1', sourceRange: 'A1:A3', chartId: 'Chart-second' },
    ]);

    // The first chart's structuralOp (A1:B3) never got an id, so revert should
    // refuse rather than silently drop that half of the batch.
    await expect(service.revert(preview.changeSetId)).rejects.toThrow(/A1:B3/);
  });
});
