import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/** Same in-memory fake Mongoose model precedent as change-set.service.chart-revert.spec.ts. */
function createInMemoryModel() {
  const store = new Map<string, Record<string, unknown>>();
  function withExec<T>(value: T) {
    const promise = Promise.resolve(value);
    (promise as Promise<T> & { exec: () => Promise<T> }).exec = () => promise;
    return promise;
  }
  return {
    create: jest.fn(async (data: Record<string, unknown>) => {
      const doc: Record<string, unknown> = { ...data, save: jest.fn(async () => {}) };
      store.set(doc.changeSetId as string, doc);
      return doc;
    }),
    findOne: jest.fn(({ changeSetId }: { changeSetId: string }) =>
      withExec(store.get(changeSetId) ?? null),
    ),
    findOneAndUpdate: jest.fn(() => withExec(null)),
  };
}

function createInMemoryConversationModel() {
  return { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
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

/**
 * TASKS.md #246 follow-up — live-tested smoke test for "Copy the Purchase
 * Register sheet and name it March Copy" after the planner fix (single
 * COPY_SHEET, no more ADD_SHEET+COPY_FILTERED_RANGE split): the reported
 * WorkbookContext for a 61-row sheet, exactly like the real one, only ever
 * carries a compressed sample (here: header + 2 data rows, `rowCount: 61`
 * claimed) — the same "TASKS.md F13" shape the client's own
 * `extendRangeToUsedRows` works around at real-apply time. Confirms the
 * preview no longer confidently predicts wrong content for the copy's
 * un-sampled rows and reports a false "N cells do not match" once Office.js
 * performs the real, correct, full copy.
 */
describe('ChangeSetService — COPY_SHEET preview does not fabricate a false diff (TASKS.md #246)', () => {
  const purchaseRegisterContext: WorkbookContext = {
    activeSheetName: 'Purchase Register',
    sheets: [
      {
        name: 'Purchase Register',
        usedRange: 'A1:B61',
        // Claims 61 real rows, but the sampled `values` below only carries 3 —
        // exactly the shape a compressed/sampled WorkbookContext takes for a
        // sheet larger than what got fully transcribed into the prompt.
        rowCount: 61,
        columnCount: 2,
        values: [
          ['Invoice No', 'GSTIN'],
          ['INV-2024-0002', '29AAAPL1234C1Z5'],
          ['INV-2024-0003', '32AAACK5678D1Z2'],
        ],
        formulas: [[], [], []],
        numberFormats: [[], [], []],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('produces no changes for the copied destination sheet — Office.js is the sole source of truth for a real copy', async () => {
    const service = buildService();
    const action: Action = {
      type: 'COPY_SHEET',
      sheetName: 'Purchase Register',
      newSheetName: 'March Copy',
    } as Action;

    const record = await service.createPreview({
      conversationId: 'conv_test',
      traceId: 'trace_test',
      prompt: "Copy the Purchase Register sheet and name it March Copy",
      context: purchaseRegisterContext,
      actions: [action],
      workbookId: 'wb_test',
    });

    const destChanges = record.changes.filter((c) => c.sheet === 'March Copy');
    expect(destChanges).toHaveLength(0);
    // The source sheet is untouched either way — belt and suspenders.
    expect(record.changes.filter((c) => c.sheet === 'Purchase Register')).toHaveLength(0);
    expect(record.irreversibleActionTypes).toContain('COPY_SHEET');
  });

  it('ADD_SHEET{copyFrom} gets the same treatment as a plain COPY_SHEET', async () => {
    const service = buildService();
    const action: Action = {
      type: 'ADD_SHEET',
      name: 'March Copy',
      copyFrom: 'Purchase Register',
    } as Action;

    const record = await service.createPreview({
      conversationId: 'conv_test2',
      traceId: 'trace_test2',
      prompt: 'Copy the Purchase Register sheet and name it March Copy',
      context: purchaseRegisterContext,
      actions: [action],
      workbookId: 'wb_test2',
    });

    expect(record.changes.filter((c) => c.sheet === 'March Copy')).toHaveLength(0);
  });
});
