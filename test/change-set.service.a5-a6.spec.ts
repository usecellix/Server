import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Same in-memory fake Mongoose model as change-set.service.e2e-dashboard-revert.spec.ts
 * (TASKS.md #20) — real create/findOne/findOneAndUpdate semantics backed by a Map, no
 * mongodb-memory-server in this repo.
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

const twoSheetContext: WorkbookContext = {
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
    {
      name: 'Sheet2',
      usedRange: 'A1:A2',
      rowCount: 2,
      columnCount: 1,
      values: [['Existing'], ['#REF!']],
      formulas: [[''], ['']],
      numberFormats: [['General'], ['General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('ChangeSetService — PRD Tier A metrics (TASKS.md #48 A5 / #49 A6)', () => {
  it('reports empty unintendedChanges/formulaErrorsIntroduced for a normal, in-scope, error-free change', async () => {
    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-a5a6-clean',
      traceId: 'trace-a5a6-clean',
      prompt: 'Update the total',
      context: twoSheetContext,
      actions: [
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 } as Action,
      ],
    });

    expect(preview.unintendedChanges).toEqual([]);
    expect(preview.formulaErrorsIntroduced).toEqual([]);
  });

  it('does not flag a multi-sheet change when every touched sheet was declared by some action (A5 negative case)', async () => {
    // Confirms the scope check is genuinely "sheet declared by some action in the
    // batch," not "only one sheet changed" — a legitimate multi-sheet batch (both
    // actions declare their own sheetName) must produce zero unintendedChanges even
    // though more than one sheet changed. The positive case — a change landing on a
    // sheet no action declared at all — is exercised at the pure-function level in
    // diff.engine.spec.ts (constructing the mismatch directly is the honest way to
    // test it: no real action type in this codebase legitimately mutates a sheet
    // other than the one it declares, so there's no non-contrived way to reach that
    // state through the real service without injecting a bug).
    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-a5-flag',
      traceId: 'trace-a5-flag',
      prompt: 'Update Sheet1 only',
      context: twoSheetContext,
      actions: [
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 } as Action,
        { type: 'SET_CELL', sheetName: 'Sheet2', row: 0, col: 0, value: 'Changed' } as Action,
      ],
    });

    // Sanity: both sheets really did change.
    expect(preview.changes.some((c) => c.sheet === 'Sheet1')).toBe(true);
    expect(preview.changes.some((c) => c.sheet === 'Sheet2')).toBe(true);
    // But both were declared (both actions carry sheetName), so nothing is unintended —
    // this confirms the scope check isn't just "did more than one sheet change."
    expect(preview.unintendedChanges).toEqual([]);
  });

  it('detects a newly introduced error and does not double-count the pre-existing one elsewhere (A6)', async () => {
    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-a6-flag',
      traceId: 'trace-a6-flag',
      prompt: 'Update the total (simulating an executor bug that wrote a literal error)',
      context: twoSheetContext,
      actions: [
        // Genuinely new: Sheet1!B2 goes from a plain number to an error string.
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: '#REF!' } as Action,
      ],
    });

    expect(preview.formulaErrorsIntroduced).toEqual([{ cell: 'B2', sheet: 'Sheet1', error: '#REF!' }]);
    // Sheet2!A2 already had '#REF!' in the seed context and this batch never touches
    // Sheet2 at all — it must never appear in changes, let alone formulaErrorsIntroduced.
    expect(preview.changes.some((c) => c.sheet === 'Sheet2')).toBe(false);
    expect(preview.formulaErrorsIntroduced.some((e) => e.sheet === 'Sheet2')).toBe(false);
  });
});
