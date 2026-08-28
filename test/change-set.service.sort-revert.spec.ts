import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { beforeStateToBulkInverseActions } from '../src/audit/diff.engine';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Same in-memory fake Mongoose model as the CONDITIONAL_FORMAT/CREATE_CHART
 * revert specs (TASKS.md #20's precedent) — real create/findOne/findOneAndUpdate
 * semantics backed by a Map, no live DB.
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

// A sparse range (a fully-blank data row inside the sort range) — the exact
// shape virtualApply.ts's virtualSortRange bails out on rather than risk a
// garbage diff from an incomplete shadow copy of the sheet.
const sparseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B4',
      rowCount: 4,
      columnCount: 2,
      values: [
        ['Item', 'Score'],
        ['A', 10],
        ['', ''],
        ['B', 20],
      ],
      formulas: [
        ['', ''],
        ['', ''],
        ['', ''],
        ['', ''],
      ],
      numberFormats: [
        ['General', 'General'],
        ['General', 'General'],
        ['General', 'General'],
        ['General', 'General'],
      ],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('ChangeSetService — SORT_RANGE revert on a sparse range (TASKS.md #93)', () => {
  it('shadow-based preview reports 0 changes for a sparse sort (reproduces the bug)', async () => {
    const action: Action = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:B4',
      key: 1,
      ascending: true,
      hasHeaders: true,
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-sort-1',
      traceId: 'trace-sort-1',
      prompt: 'sort by score',
      context: sparseContext,
      actions: [action],
    });

    // The shadow's virtualSortRange refused to simulate the sparse range —
    // before === after, so nothing shows up as changed.
    expect(preview.changes).toEqual([]);
  });

  it('accepting frontend-reported real changes at apply time gives Revert something to actually undo', async () => {
    const action: Action = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:B4',
      key: 1,
      ascending: true,
      hasHeaders: true,
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-sort-2',
      traceId: 'trace-sort-2',
      prompt: 'sort by score',
      context: sparseContext,
      actions: [action],
    });
    expect(preview.changes).toEqual([]);

    // What the frontend actually saw, reading the real sheet before/after
    // performing the sort for real via Office.js.
    const frontendChanges = [
      { cell: 'A2', sheet: 'Sheet1', before: 'A', after: 'B', isHardcoded: true },
      { cell: 'B2', sheet: 'Sheet1', before: 10, after: 20, isHardcoded: true },
      { cell: 'A3', sheet: 'Sheet1', before: 'B', after: 'A', isHardcoded: true },
      { cell: 'B3', sheet: 'Sheet1', before: 20, after: 10, isHardcoded: true },
    ];

    await service.markApplied(preview.changeSetId, undefined, undefined, frontendChanges);

    const { changeSet, inverseActions } = await service.revert(preview.changeSetId);

    expect(changeSet.changes).toHaveLength(4);
    // TASKS.md #100 — one bulk SET_RANGE_VALUES per sheet instead of one
    // SET_CELL per touched cell.
    expect(inverseActions).toHaveLength(1);
    expect(inverseActions[0]).toMatchObject({ type: 'SET_RANGE_VALUES', sheetName: 'Sheet1' });
    const operations = (inverseActions[0] as unknown as { operations: { address: string; value: unknown }[] })
      .operations;
    expect(operations).toEqual(
      expect.arrayContaining([
        { address: 'A2', value: 'A' },
        { address: 'B2', value: 10 },
        { address: 'A3', value: 'B' },
        { address: 'B3', value: 20 },
      ]),
    );
  });

  it('merges frontend-reported changes with any pre-existing shadow-computed changes rather than dropping them', async () => {
    // A batch with a SET_CELL (shadow captures it fine) alongside the sparse
    // SORT_RANGE (shadow can't). Both must survive into the final change set.
    const actions: Action[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', address: 'C1', value: 'Note' } as Action,
      {
        type: 'SORT_RANGE',
        sheetName: 'Sheet1',
        range: 'A1:B4',
        key: 1,
        ascending: true,
        hasHeaders: true,
      } as Action,
    ];

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-sort-3',
      traceId: 'trace-sort-3',
      prompt: 'add a note and sort by score',
      context: sparseContext,
      actions,
    });

    // The SET_CELL half of the batch is captured by the shadow diff normally.
    expect(preview.changes).toEqual([
      expect.objectContaining({ cell: 'C1', sheet: 'Sheet1', after: 'Note' }),
    ]);

    const frontendChanges = [
      { cell: 'A2', sheet: 'Sheet1', before: 'A', after: 'B', isHardcoded: true },
    ];
    await service.markApplied(preview.changeSetId, undefined, undefined, frontendChanges);

    const { changeSet } = await service.revert(preview.changeSetId);
    const cells = changeSet.changes.map((c) => c.cell).sort();
    expect(cells).toEqual(['A2', 'C1']);
  });
});

// A fully-populated (non-sparse) range — virtualSortRange does NOT bail out
// here, it actually simulates the sort. This is the shape that surfaced the
// real bug: beforeState only ever contains the CHANGED cells, not a full-range
// snapshot, so replaying the sort forward on that partial reconstruction (the
// old self-verification's job) reorders garbage in the cells beforeState never
// captured and reports them as "would not converge" — even though the actual
// inverse actions (absolute writes of real captured values) are already correct.
const nonSparseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B4',
      rowCount: 4,
      columnCount: 2,
      values: [
        ['Item', 'Score'],
        ['A', 30],
        ['B', 10],
        ['C', 20],
      ],
      formulas: [
        ['', ''],
        ['', ''],
        ['', ''],
        ['', ''],
      ],
      numberFormats: [
        ['General', 'General'],
        ['General', 'General'],
        ['General', 'General'],
        ['General', 'General'],
      ],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('ChangeSetService — SORT_RANGE revert on a fully-populated range (TASKS.md #99)', () => {
  it('reverts successfully even though beforeState only covers the changed cells, not the full range', async () => {
    const action: Action = {
      type: 'SORT_RANGE',
      sheetName: 'Sheet1',
      range: 'A1:B4',
      key: 1,
      ascending: true,
      hasHeaders: true,
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-sort-4',
      traceId: 'trace-sort-4',
      prompt: 'sort by score',
      context: nonSparseContext,
      actions: [action],
    });

    // Real ascending-by-score result: B(10), C(20), A(30) — every data row moved.
    const frontendChanges = [
      { cell: 'A2', sheet: 'Sheet1', before: 'A', after: 'B', isHardcoded: true },
      { cell: 'B2', sheet: 'Sheet1', before: 30, after: 10, isHardcoded: true },
      { cell: 'A3', sheet: 'Sheet1', before: 'B', after: 'C', isHardcoded: true },
      { cell: 'B3', sheet: 'Sheet1', before: 10, after: 20, isHardcoded: true },
      { cell: 'A4', sheet: 'Sheet1', before: 'C', after: 'A', isHardcoded: true },
      { cell: 'B4', sheet: 'Sheet1', before: 20, after: 30, isHardcoded: true },
    ];

    await service.markApplied(preview.changeSetId, undefined, undefined, frontendChanges);

    // Must not throw RevertVerificationError — this is the exact failure the
    // user hit live ("52 cell(s) would not converge") before this fix.
    const { changeSet, inverseActions } = await service.revert(preview.changeSetId);

    expect(changeSet.status).toBe('reverted');
    // TASKS.md #100 — one bulk SET_RANGE_VALUES per sheet instead of one
    // SET_CELL per touched cell (this used to assert 6 separate SET_CELLs).
    expect(inverseActions).toHaveLength(1);
    expect(inverseActions[0]).toMatchObject({
      type: 'SET_RANGE_VALUES',
      sheetName: 'Sheet1',
      range: 'A2:B4',
    });
    const operations = (inverseActions[0] as unknown as { operations: { address: string; value: unknown }[] })
      .operations;
    expect(operations).toEqual(
      expect.arrayContaining([
        { address: 'A2', value: 'A' },
        { address: 'B2', value: 30 },
        { address: 'A3', value: 'B' },
        { address: 'B3', value: 10 },
        { address: 'A4', value: 'C' },
        { address: 'B4', value: 20 },
      ]),
    );
  });

  it('still runs full self-verification for change sets with no frontend-reported changes (regression guard)', async () => {
    // A normal SET_CELL-only batch (shadow diff is fully accurate) must still
    // go through the real convergence check — this flag must not weaken
    // verification for change sets that never needed the escape hatch.
    const action: Action = {
      type: 'SET_CELL',
      sheetName: 'Sheet1',
      address: 'C1',
      value: 'Note',
    } as Action;

    const service = buildService();
    const preview = await service.createPreview({
      conversationId: 'conv-sort-5',
      traceId: 'trace-sort-5',
      prompt: 'add a note',
      context: nonSparseContext,
      actions: [action],
    });

    await service.markApplied(preview.changeSetId);
    const { changeSet } = await service.revert(preview.changeSetId);
    expect(changeSet.status).toBe('reverted');
  });
});

describe('beforeStateToBulkInverseActions (TASKS.md #100)', () => {
  it('declines (returns null) if any touched cell captured a real formula — bulk write cannot restore formulas', () => {
    const beforeState = {
      'Sheet1!A2': { value: 10, formula: '', format: 'General' },
      'Sheet1!A3': { value: 5, formula: '=A2*2', format: 'General' },
    };
    const changes = [
      { cell: 'A2', sheet: 'Sheet1', before: 10, after: 20, isHardcoded: true },
      { cell: 'A3', sheet: 'Sheet1', before: 5, after: 10, isHardcoded: false },
    ];
    expect(beforeStateToBulkInverseActions(beforeState, changes)).toBeNull();
  });

  it('groups touched cells by sheet into one bounding-box action per sheet', () => {
    const beforeState = {
      'Sheet1!A2': { value: 'x', formula: '', format: 'General' },
      'Sheet2!B5': { value: 'y', formula: '', format: 'General' },
    };
    const changes = [
      { cell: 'A2', sheet: 'Sheet1', before: 'x', after: 'z', isHardcoded: true },
      { cell: 'B5', sheet: 'Sheet2', before: 'y', after: 'w', isHardcoded: true },
    ];
    const actions = beforeStateToBulkInverseActions(beforeState, changes);
    expect(actions).toHaveLength(2);
    expect(actions).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: 'SET_RANGE_VALUES', sheetName: 'Sheet1', range: 'A2:A2' }),
        expect.objectContaining({ type: 'SET_RANGE_VALUES', sheetName: 'Sheet2', range: 'B5:B5' }),
      ]),
    );
  });
});
