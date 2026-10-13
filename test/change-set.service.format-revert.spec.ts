import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { RevertNoOpError } from '../src/audit/errors/revert-noop.error';
import { FormatSnapshot } from '../src/audit/format-snapshot';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #400 — a formatting request can be reverted. The add-in reads the
 * previous format off Excel just before applying and reports it with the apply
 * call; revert turns it back into actions. Same in-memory model as the chart and
 * conditional-format revert specs.
 */
function createInMemoryModel() {
  const store = new Map<string, Record<string, unknown>>();
  const withExec = <T>(value: T) => {
    const promise = Promise.resolve(value);
    (promise as Promise<T> & { exec: () => Promise<T> }).exec = () => promise;
    return promise;
  };
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
    findOne: jest.fn(({ changeSetId }: { changeSetId: string }) => withExec(store.get(changeSetId) ?? null)),
    findOneAndUpdate: jest.fn(
      ({ changeSetId, status }: { changeSetId: string; status: string }, update: Record<string, unknown>) => {
        const doc = store.get(changeSetId);
        if (!doc || doc.status !== status) return withExec(null);
        Object.assign(doc, update);
        return withExec(doc);
      },
    ),
  };
}

function buildService() {
  return new ChangeSetService(
    createInMemoryModel() as unknown as Model<ChangeSetDocument>,
    { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) } as never,
    { appendTerminalByChangeSet: jest.fn(), appendTerminalByConversationId: jest.fn() } as unknown as WorkflowTraceService,
  );
}

const context: WorkbookContext = {
  activeSheetName: 'Monthly Summary',
  sheets: [
    {
      name: 'Monthly Summary',
      usedRange: 'A3:D5',
      rowCount: 5,
      columnCount: 4,
      values: [['', '', '', ''], ['', '', '', ''], ['Month', 'Debits', 'Credits', 'Net'], ['Apr', 1, 2, 1], ['May', 3, 4, 1]],
      formulas: [['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']],
      numberFormats: Array.from({ length: 5 }, () => ['General', 'General', 'General', 'General']),
      structure: 'data_table',
      headerRowIndex: 2,
    },
  ],
  namedRanges: [],
  tables: [],
};

const headerBand = {
  type: 'FORMAT_RANGE',
  sheetName: 'Monthly Summary',
  row: 2,
  col: 0,
  rowCount: 1,
  colCount: 4,
  format: { bold: true, fillColor: '#4472C4', fontColor: '#FFFFFF' },
} as Action;
const amounts = {
  type: 'FORMAT_RANGE',
  sheetName: 'Monthly Summary',
  row: 3,
  col: 1,
  rowCount: 2,
  colCount: 3,
  format: { numberFormat: '#,##0.00' },
} as Action;
const autofit = { type: 'AUTOFIT_COLUMNS', sheetName: 'Monthly Summary', col: 0, colCount: 4 } as Action;

const before: FormatSnapshot[] = [
  {
    kind: 'format',
    sheetName: 'Monthly Summary',
    restorable: true,
    row: 2,
    col: 0,
    rowCount: 1,
    colCount: 4,
    palette: [{ bold: false, clearFill: true, fontColor: '#000000' }],
    grid: [[0, 0, 0, 0]],
  },
  {
    kind: 'format',
    sheetName: 'Monthly Summary',
    restorable: true,
    row: 3,
    col: 1,
    rowCount: 2,
    colCount: 3,
    palette: [{ numberFormat: 'General' }],
    grid: [[0, 0, 0], [0, 0, 0]],
  },
  { kind: 'columns', sheetName: 'Monthly Summary', restorable: true, widths: [{ col: 0, width: 64 }, { col: 1, width: 64 }, { col: 2, width: 64 }, { col: 3, width: 64 }] },
];

async function preview(service: ChangeSetService, actions: Action[]) {
  return service.createPreview({
    conversationId: 'conv-format',
    traceId: 'trace-format',
    prompt: 'format the summary',
    context,
    actions,
  });
}

describe('ChangeSetService — formatting preview → apply → revert (TASKS.md #400)', () => {
  it('no longer flags an ordinary formatting request as irreversible', async () => {
    const result = await preview(buildService(), [headerBand, amounts, autofit]);
    expect(result.irreversibleActionTypes).toEqual([]);
  });

  it('puts the previous formatting and column widths back, newest action first', async () => {
    const service = buildService();
    const created = await preview(service, [headerBand, amounts, autofit]);
    await service.markApplied(created.changeSetId, undefined, undefined, undefined, before);

    const { inverseActions } = await service.revert(created.changeSetId);

    expect(inverseActions).toEqual([
      { type: 'SET_COLUMN_WIDTH', sheetName: 'Monthly Summary', col: 0, colCount: 4, width: 64 },
      { type: 'FORMAT_RANGE', sheetName: 'Monthly Summary', row: 3, col: 1, rowCount: 2, colCount: 3, format: { numberFormat: 'General' } },
      {
        type: 'FORMAT_RANGE',
        sheetName: 'Monthly Summary',
        row: 2,
        col: 0,
        rowCount: 1,
        colCount: 4,
        format: { bold: false, clearFill: true, fontColor: '#000000' },
      },
    ]);
  });

  it('refuses to revert when the apply call never carried the snapshots (fails closed, never a false success)', async () => {
    const service = buildService();
    const created = await preview(service, [headerBand]);
    await service.markApplied(created.changeSetId); // add-in reported nothing

    await expect(service.revert(created.changeSetId)).rejects.toBeInstanceOf(RevertNoOpError);
  });

  it('refuses to revert when one of the snapshots could not be taken', async () => {
    const service = buildService();
    const created = await preview(service, [headerBand]);
    await service.markApplied(created.changeSetId, undefined, undefined, undefined, [
      { kind: 'format', sheetName: 'Monthly Summary', restorable: false, reason: 'range too large' },
    ]);

    await expect(service.revert(created.changeSetId)).rejects.toBeInstanceOf(RevertNoOpError);
  });

  it('refuses to revert when the snapshots are fewer than the formatting actions', async () => {
    const service = buildService();
    const created = await preview(service, [headerBand, amounts]);
    await service.markApplied(created.changeSetId, undefined, undefined, undefined, [before[0]]);

    await expect(service.revert(created.changeSetId)).rejects.toBeInstanceOf(RevertNoOpError);
  });

  it('does not store a snapshot whose shape is not what it claims', async () => {
    const service = buildService();
    const created = await preview(service, [headerBand]);
    await service.markApplied(created.changeSetId, undefined, undefined, undefined, [
      { ...before[0], grid: [[0, 9, 0, 0]] },
    ]);

    await expect(service.revert(created.changeSetId)).rejects.toBeInstanceOf(RevertNoOpError);
  });

  it('keeps borders irreversible and says so before Accept', async () => {
    const result = await preview(buildService(), [
      { ...headerBand, format: { borders: 'all' } } as unknown as Action,
    ]);
    expect(result.irreversibleActionTypes).toEqual(['FORMAT_RANGE']);
  });

  it('keeps a range too large to snapshot irreversible', async () => {
    const result = await preview(buildService(), [
      { ...headerBand, rowCount: 1_000_000, colCount: 16 } as Action,
    ]);
    expect(result.irreversibleActionTypes).toEqual(['FORMAT_RANGE']);
  });
});
