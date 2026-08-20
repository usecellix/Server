import { ChangeSetService } from '../src/audit/change-set.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';

/**
 * TASKS.md #24 — createPreview() resolves ChangeSet.workbookId from the
 * originating conversation by conversationId (see ChangeSetService.resolveWorkbookId),
 * rather than requiring every one of its several call sites in conversation.service.ts
 * to remember to thread it through. This suite exercises that resolution directly.
 */

const workflowTraceStub = () =>
  ({
    appendTerminalByChangeSet: jest.fn(),
    appendTerminalByConversationId: jest.fn(),
  }) as unknown as WorkflowTraceService;

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:A1',
      rowCount: 1,
      columnCount: 1,
      values: [['A']],
      formulas: [['']],
      numberFormats: [['General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

const actions: Action[] = [{ type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'x' }];

function buildService(conversationDoc: { workbookId?: string } | null) {
  const created: { workbookId?: string } = {};
  const changeSetModel = {
    create: jest.fn(async (doc: { workbookId?: string }) => {
      created.workbookId = doc.workbookId;
      return { ...doc, status: 'previewed', timestamp: new Date() };
    }),
  };
  const findOne = jest.fn(() => ({
    lean: () => ({ exec: () => Promise.resolve(conversationDoc) }),
  }));
  const conversationModel = { findOne };
  const service = new ChangeSetService(
    changeSetModel as never,
    conversationModel as never,
    workflowTraceStub(),
  );
  return { service, created, findOne };
}

describe('ChangeSetService.createPreview — workbookId resolution (TASKS.md #24)', () => {
  it('resolves workbookId from the originating conversation and stores it on the change set', async () => {
    const { service, created, findOne } = buildService({ workbookId: 'wb_shared-1' });

    const record = await service.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'set A1',
      context,
      actions,
    });

    expect(findOne).toHaveBeenCalledWith({ conversationId: 'conv-1' }, { workbookId: 1 });
    expect(created.workbookId).toBe('wb_shared-1');
    expect(record.workbookId).toBe('wb_shared-1');
  });

  it('leaves workbookId entirely unset when the conversation has none (backward compatible)', async () => {
    const { service, created } = buildService({ workbookId: undefined });

    const record = await service.createPreview({
      conversationId: 'conv-2',
      traceId: 'trace-1',
      prompt: 'set A1',
      context,
      actions,
    });

    expect(created.workbookId).toBeUndefined();
    expect(record.workbookId).toBeUndefined();
  });

  it('leaves workbookId unset when the conversation cannot be found (expired/legacy)', async () => {
    const { service, created } = buildService(null);

    const record = await service.createPreview({
      conversationId: 'conv-missing',
      traceId: 'trace-1',
      prompt: 'set A1',
      context,
      actions,
    });

    expect(created.workbookId).toBeUndefined();
    expect(record.workbookId).toBeUndefined();
  });

  it('an explicit input.workbookId overrides the conversation lookup (used by tests/future callers)', async () => {
    const { service, created, findOne } = buildService({ workbookId: 'wb_from-conversation' });

    await service.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'set A1',
      context,
      actions,
      workbookId: 'wb_explicit',
    });

    expect(created.workbookId).toBe('wb_explicit');
    expect(findOne).not.toHaveBeenCalled();
  });

  it('does not throw and leaves workbookId unset if the conversation lookup itself fails', async () => {
    const changeSetModel = {
      create: jest.fn(async (doc: { workbookId?: string }) => ({
        ...doc,
        status: 'previewed',
        timestamp: new Date(),
      })),
    };
    const conversationModel = {
      findOne: jest.fn(() => ({
        lean: () => ({ exec: () => Promise.reject(new Error('mock db error')) }),
      })),
    };
    const service = new ChangeSetService(
      changeSetModel as never,
      conversationModel as never,
      workflowTraceStub(),
    );

    const record = await service.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'set A1',
      context,
      actions,
    });

    expect(record.workbookId).toBeUndefined();
  });
});
