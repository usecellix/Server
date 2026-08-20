import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import { diffShadowsFully, snapshotBeforeState } from '../src/audit/diff.engine';

/**
 * In-memory fake Mongoose model — real `create`/`findOne`/`findOneAndUpdate` semantics,
 * backed by a Map instead of a live DB (no mongodb-memory-server in this repo, see
 * TASKS.md #11/#19's precedent for the same constraint). Good enough to exercise
 * ChangeSetService's real create→apply→revert pipeline end to end, not just diff.engine's
 * standalone functions.
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

const salesContext: WorkbookContext = {
  activeSheetName: 'Sales',
  sheets: [
    {
      name: 'Sales',
      usedRange: 'A1:C3',
      rowCount: 3,
      columnCount: 3,
      values: [
        ['Region', 'Qty', 'Price'],
        ['North', 10, 2],
        ['South', 5, 3],
      ],
      formulas: [
        ['', '', ''],
        ['', '', ''],
        ['', '', ''],
      ],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('ChangeSetService — end-to-end dashboard-style build → single revert (TASKS.md #20)', () => {
  it(
    'reverts a multi-piece build (new sheet, formulas, inserted column, table, merge) back to a workbook ' +
      'indistinguishable from the pre-request state',
    async () => {
      // Scoping note (see TASKS.md #20's own row for the full explanation): the task's
      // literal PRD wording asks for a chart and a conditional-format rule too, but
      // CREATE_CHART (#15) and conditional formatting (M3) have no real inverse today —
      // #15 is deliberately deferred, M3 isn't built at all. This test instead chains
      // together every structural type #12–17 actually made revertible, which is the
      // real thing this task can prove: that the *pieces*, applied together in one
      // change set, revert cleanly as a whole — not just individually, as each was
      // tested in isolation in #12–17's own spec files.
      const actions: Action[] = [
        // New "Total" column on Sales, with a formula. Targeted via afterColumn (not
        // plain afterLastColumn) so it exercises the same insert-and-shift path #13's
        // own P0 repro covers — appending past the current bounds is a different,
        // non-structural code path (see TASKS.md #13's own note on this distinction).
        {
          type: 'INSERT_COLUMN',
          sheetName: 'Sales',
          columnName: 'Total',
          afterColumn: 'Price',
          formula: '=B{row}*C{row}',
        } as Action,
        // Wrap the Sales data as a real Table
        { type: 'CREATE_TABLE', sheetName: 'Sales', range: 'A1:C3', tableName: 'SalesTable', hasHeaders: true } as Action,
        // A new Dashboard sheet summarizing Sales
        { type: 'ADD_SHEET', name: 'Dashboard' } as Action,
        { type: 'SET_FORMULA', sheetName: 'Dashboard', row: 0, col: 0, formula: '=SUM(Sales!B2:B3)' } as Action,
        { type: 'SET_CELL', sheetName: 'Dashboard', row: 1, col: 0, value: 'Summary' } as Action,
        // A merged title band on the Dashboard
        { type: 'MERGE_CELLS', sheetName: 'Dashboard', range: 'A1:B1' } as Action,
      ];

      const service = buildService();

      const preview = await service.createPreview({
        conversationId: 'conv-e2e',
        traceId: 'trace-e2e',
        prompt: 'Build a Sales dashboard',
        context: salesContext,
        actions,
      });

      // Sanity: this build actually touches every structural type under test.
      const structuralOpTypes = new Set(preview.structuralOps.map((op) => op.opType));
      expect(structuralOpTypes).toEqual(
        new Set(['INSERT_COLUMN', 'CREATE_TABLE', 'ADD_SHEET', 'MERGE_CELLS']),
      );
      expect(preview.irreversibleActionTypes).toEqual([]);

      await service.markApplied(preview.changeSetId);
      const { inverseActions } = await service.revert(preview.changeSetId);

      // Independently confirm convergence: replay the same forward actions, then the
      // returned inverse actions, and diff the result against the pristine original —
      // the same technique #19's fail-closed check uses internally, applied here from
      // the outside as an independent check on the real service's output.
      const original = buildShadowWorkbook(salesContext);
      const afterForward = virtualApply(original, actions);
      const afterRevert = virtualApply(afterForward, inverseActions);
      const residualDiff = diffShadowsFully(original, afterRevert);

      expect(residualDiff).toEqual([]);
      expect(afterRevert.sheets.has('Dashboard')).toBe(false);
      expect(afterRevert.sheets.get('Sales')?.columnCount).toBe(
        original.sheets.get('Sales')?.columnCount,
      );
      expect(snapshotBeforeState(afterRevert)).toEqual(snapshotBeforeState(original));
    },
  );
});
