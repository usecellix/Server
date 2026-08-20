import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { diffShadowsFully } from '../src/audit/diff.engine';

/**
 * Same in-memory fake Mongoose model as change-set.service.e2e-dashboard-revert.spec.ts
 * (TASKS.md #20's precedent) — real create/findOne/findOneAndUpdate semantics backed by a
 * Map, no live DB.
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

describe('ChangeSetService — CONDITIONAL_FORMAT create → apply → revert (TASKS.md #40)', () => {
  it('patches the real ruleId in at markApplied time and builds a working DELETE_CONDITIONAL_FORMAT revert', async () => {
    const action: Action = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Sheet1',
      range: 'B2:B3',
      rule: { kind: 'cellValue', operator: 'greaterThan', value: 15, format: { fillColor: '#FFC7CE' } },
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-cf',
      traceId: 'trace-cf',
      prompt: 'highlight scores above 15',
      context: scoresContext,
      actions: [action],
    });

    // Preview time: the structural op exists but has no ruleId yet.
    expect(preview.structuralOps).toEqual([
      expect.objectContaining({ opType: 'CONDITIONAL_FORMAT', sheetName: 'Sheet1', params: { range: 'B2:B3' } }),
    ]);
    expect(preview.irreversibleActionTypes).toEqual([]);

    // Apply-time: the frontend reports back the real Excel-assigned id.
    await service.markApplied(preview.changeSetId, [
      { sheetName: 'Sheet1', range: 'B2:B3', ruleId: 'cf-real-42' },
    ]);

    const { inverseActions } = await service.revert(preview.changeSetId);
    expect(inverseActions).toContainEqual({
      type: 'DELETE_CONDITIONAL_FORMAT',
      sheetName: 'Sheet1',
      ruleId: 'cf-real-42',
    });

    // Independent convergence check, same technique as the #20 e2e test.
    const original = buildShadowWorkbook(scoresContext);
    const afterForward = virtualApply(original, [action]);
    const afterRevert = virtualApply(afterForward, inverseActions);
    expect(diffShadowsFully(original, afterRevert)).toEqual([]);
  });

  it('leaves the change set unrevertable if markApplied is never given the created id (fails closed)', async () => {
    const action: Action = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Sheet1',
      range: 'B2:B3',
      rule: { kind: 'cellValue', operator: 'greaterThan', value: 15, format: { fillColor: '#FFC7CE' } },
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-cf-2',
      traceId: 'trace-cf-2',
      prompt: 'highlight scores above 15',
      context: scoresContext,
      actions: [action],
    });

    await service.markApplied(preview.changeSetId); // no createdConditionalFormatIds
    await expect(service.revert(preview.changeSetId)).rejects.toThrow(/no ruleId was ever reported/);
  });

  it('does not touch structuralOps for a MODIFY (existingRuleId set) — nothing to patch, nothing captured', async () => {
    const action: Action = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Sheet1',
      range: 'B2:B3',
      existingRuleId: 'cf-already-exists',
      rule: { kind: 'cellValue', operator: 'greaterThan', value: 25, format: { fillColor: '#FFC7CE' } },
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-cf-3',
      traceId: 'trace-cf-3',
      prompt: 'change the threshold to 25',
      context: scoresContext,
      actions: [action],
    });

    expect(preview.structuralOps).toEqual([]);
    expect(preview.irreversibleActionTypes).toEqual(['CONDITIONAL_FORMAT']);
  });

  it('only patches a structuralOp whose sheetName+range matches the reported id, in a batch with two creates', async () => {
    const actions: Action[] = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B3',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 15, format: { fillColor: '#FFC7CE' } },
      } as Action,
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'A2:A3',
        rule: { kind: 'cellValue', operator: 'lessThan', value: 5, format: { fillColor: '#C6EFCE' } },
      } as Action,
    ];

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-cf-4',
      traceId: 'trace-cf-4',
      prompt: 'highlight two columns',
      context: scoresContext,
      actions,
    });
    expect(preview.structuralOps).toHaveLength(2);

    // Only report the id for the SECOND rule (A2:A3) this time.
    await service.markApplied(preview.changeSetId, [
      { sheetName: 'Sheet1', range: 'A2:A3', ruleId: 'cf-second' },
    ]);

    // The first rule's structuralOp (B2:B3) never got an id, so revert should
    // refuse rather than silently drop that half of the batch.
    await expect(service.revert(preview.changeSetId)).rejects.toThrow(/B2:B3/);
  });
});
