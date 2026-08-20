import { Model } from 'mongoose';
import { ChangeSetService } from '../src/audit/change-set.service';
import { CheckpointService } from '../src/audit/checkpoint.service';
import { RestoreVerificationError } from '../src/audit/errors/restore-verification.error';
import { ChangeSetDocument } from '../src/audit/schemas/change-set.schema';
import { CheckpointDocument } from '../src/audit/schemas/checkpoint.schema';
import { WorkflowTraceService } from '../src/common/logging/workflow-trace.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #26-30 — M2 (Checkpoints). In-memory fake Mongoose models, same
 * pattern/precedent as change-set.service.e2e-dashboard-revert.spec.ts (#20):
 * real create/findOne/find/findOneAndUpdate semantics backed by a Map, good
 * enough to exercise ChangeSetService + CheckpointService's real pipeline
 * (auto-checkpoint trigger -> apply -> restore) rather than only diff.engine's
 * standalone functions.
 */

type Sortable = Record<string, unknown> & { timestamp?: Date; createdAt?: Date; _id?: number };

function withExec<T>(value: T) {
  const promise = Promise.resolve(value);
  (promise as Promise<T> & { exec: () => Promise<T> }).exec = () => promise;
  return promise;
}

/**
 * Sorts by `_id` (a plain incrementing counter standing in for MongoDB's real
 * ObjectId) when the requested field is `_id`, otherwise by the named Date
 * field. Mirrors production's switch to `_id` ordering (checkpoint.service.ts):
 * two real change sets created in rapid succession can share the same
 * millisecond-resolution `timestamp`, but ObjectIds are strictly monotonic.
 */
function sortDesc<T extends Sortable>(docs: T[], field: string): T[] {
  return [...docs].sort((a, b) => {
    if (field === '_id') return (b._id ?? 0) - (a._id ?? 0);
    const av = (a[field] as Date)?.getTime?.() ?? 0;
    const bv = (b[field] as Date)?.getTime?.() ?? 0;
    return bv - av;
  });
}

function createChangeSetModel() {
  const store = new Map<string, Record<string, unknown> & { save: jest.Mock; _id: number }>();
  let seq = 0;

  const model = {
    create: jest.fn(async (data: Record<string, unknown>) => {
      const doc: Record<string, unknown> & { save: jest.Mock; _id: number } = {
        ...data,
        _id: ++seq,
        save: jest.fn(async function (this: Record<string, unknown>) {
          store.set(this.changeSetId as string, this as typeof doc);
        }),
      };
      store.set(doc.changeSetId as string, doc);
      return doc;
    }),
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
    findOne: jest.fn((filter: Record<string, unknown>) => {
      let candidates = [...store.values()];
      if (filter.changeSetId) candidates = candidates.filter((d) => d.changeSetId === filter.changeSetId);
      if (filter.workbookId) candidates = candidates.filter((d) => d.workbookId === filter.workbookId);
      if (filter.status) candidates = candidates.filter((d) => d.status === filter.status);

      let sorted = candidates;
      const chain = {
        sort: (spec: Record<string, number>) => {
          const [field] = Object.keys(spec);
          sorted = sortDesc(sorted as Sortable[], field) as typeof candidates;
          return chain;
        },
        lean: () => chain,
        exec: () => Promise.resolve(sorted[0] ?? null),
      };
      return chain;
    }),
    find: jest.fn((filter: Record<string, unknown>) => {
      let candidates = [...store.values()];
      if (filter.workbookId) candidates = candidates.filter((d) => d.workbookId === filter.workbookId);
      if (filter.status) candidates = candidates.filter((d) => d.status === filter.status);
      const idFilter = filter._id as { $gt?: number } | undefined;
      if (idFilter?.$gt !== undefined) {
        candidates = candidates.filter((d) => (d._id as number) > idFilter.$gt!);
      }

      let sorted = candidates;
      const chain = {
        sort: (spec: Record<string, number>) => {
          const [field] = Object.keys(spec);
          sorted = sortDesc(sorted as Sortable[], field) as typeof candidates;
          return chain;
        },
        exec: () => Promise.resolve(sorted),
      };
      return chain;
    }),
  };
  return model;
}

function createCheckpointModel() {
  const store = new Map<string, Record<string, unknown> & { save: jest.Mock }>();
  return {
    create: jest.fn(async (data: Record<string, unknown>) => {
      const doc: Record<string, unknown> & { save: jest.Mock } = {
        ...data,
        save: jest.fn(async function (this: Record<string, unknown>) {
          store.set(this.checkpointId as string, this as typeof doc);
        }),
      };
      store.set(doc.checkpointId as string, doc);
      return doc;
    }),
    findOne: jest.fn(({ checkpointId }: { checkpointId: string }) =>
      withExec(store.get(checkpointId) ?? null),
    ),
    find: jest.fn(({ workbookId }: { workbookId: string }) => {
      const results = [...store.values()].filter((d) => d.workbookId === workbookId);
      const chain = {
        sort: () => chain,
        limit: () => chain,
        exec: () => Promise.resolve(sortDesc(results as Sortable[], 'createdAt')),
      };
      return chain;
    }),
  };
}

function workflowTraceStub() {
  return {
    startTrace: jest.fn(),
    appendNode: jest.fn(),
    finalize: jest.fn(),
    appendTerminalByChangeSet: jest.fn(),
    appendTerminalByConversationId: jest.fn(),
  } as unknown as WorkflowTraceService;
}

function buildContext(values: unknown[][]): WorkbookContext {
  return {
    activeSheetName: 'Sales',
    sheets: [
      {
        name: 'Sales',
        usedRange: `A1:A${values.length}`,
        rowCount: values.length,
        columnCount: 1,
        values,
        formulas: values.map(() => ['']),
        numberFormats: values.map(() => ['General']),
        structure: 'data_table',
        headerRowIndex: -1,
      },
    ],
    namedRanges: [],
    tables: [],
  };
}

function setCell(row: number, value: unknown): Action {
  return { type: 'SET_CELL', sheetName: 'Sales', row, col: 0, value } as Action;
}

describe('CheckpointService (TASKS.md #26-30)', () => {
  it('#28 — creates a manual checkpoint anchored at the latest applied change set for the workbook', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    const record = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'set A1',
      context: buildContext([['orig0']]),
      actions: [setCell(0, 'new0')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(record.changeSetId);

    const checkpoint = await checkpointService.createManual({
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'before risky edit',
    });

    expect(checkpoint.trigger).toBe('manual');
    expect(checkpoint.anchorChangeSetId).toBe(record.changeSetId);
    expect(checkpoint.status).toBe('active');

    const list = await checkpointService.listByWorkbook('wb-1');
    expect(list).toHaveLength(1);
    expect(list[0].checkpointId).toBe(checkpoint.checkpointId);
  });

  it('#28 — a manual checkpoint on a workbook with no applied change sets yet anchors at "" (beginning of history)', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      workflowTraceStub(),
    );

    const checkpoint = await checkpointService.createManual({
      workbookId: 'wb-empty',
      conversationId: 'conv-1',
    });

    expect(checkpoint.anchorChangeSetId).toBe('');
  });

  it('#27 — auto-checkpoint fires when createPreview includes a destructive action type', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'delete the sheet',
      context: buildContext([['orig0']]),
      actions: [{ type: 'DELETE_SHEET', sheetName: 'Sales' } as Action],
      workbookId: 'wb-1',
    });

    const list = await checkpointService.listByWorkbook('wb-1');
    expect(list).toHaveLength(1);
    expect(list[0].trigger).toBe('auto');
    expect(list[0].label).toContain('DELETE_SHEET');
  });

  it('#27 — no auto-checkpoint for a purely non-destructive change set', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'set A1',
      context: buildContext([['orig0']]),
      actions: [setCell(0, 'new0')],
      workbookId: 'wb-1',
    });

    expect(await checkpointService.listByWorkbook('wb-1')).toHaveLength(0);
  });

  it('#27 — auto-checkpoint is skipped (not thrown) when no workbookId is resolvable', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    const record = await changeSetService.createPreview({
      conversationId: 'conv-no-workbook',
      traceId: 'trace-1',
      prompt: 'delete the sheet',
      context: buildContext([['orig0']]),
      actions: [{ type: 'DELETE_SHEET', sheetName: 'Sales' } as Action],
    });

    expect(record).toBeDefined();
    expect(checkpointModel.create).not.toHaveBeenCalled();
  });

  it('#29 — restore: three sequential changes, checkpoint after change 1, restore undoes changes 2 and 3 and leaves change 1', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    // Change 1: row 0 orig0 -> new0
    const cs1 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'change 1',
      context: buildContext([['orig0'], ['orig1'], ['orig2']]),
      actions: [setCell(0, 'new0')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs1.changeSetId);

    const checkpoint = await checkpointService.createManual({
      workbookId: 'wb-1',
      conversationId: 'conv-1',
      label: 'after change 1',
    });

    // Change 2: row 1 orig1 -> new1 (context now reflects change 1's effect)
    const cs2 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-2',
      prompt: 'change 2',
      context: buildContext([['new0'], ['orig1'], ['orig2']]),
      actions: [setCell(1, 'new1')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs2.changeSetId);

    // Change 3: row 2 orig2 -> new2
    const cs3 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-3',
      prompt: 'change 3',
      context: buildContext([['new0'], ['new1'], ['orig2']]),
      actions: [setCell(2, 'new2')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs3.changeSetId);

    const result = await checkpointService.restore(checkpoint.checkpointId);

    expect(result.revertedChangeSetIds.sort()).toEqual([cs2.changeSetId, cs3.changeSetId].sort());
    expect(result.checkpoint.status).toBe('restored');
    expect(result.inverseActions.length).toBeGreaterThan(0);

    // The reverted change sets are marked as such; change 1 remains applied.
    const cs1After = await changeSetService.getById(cs1.changeSetId);
    const cs2After = await changeSetService.getById(cs2.changeSetId);
    const cs3After = await changeSetService.getById(cs3.changeSetId);
    expect(cs1After?.status).toBe('applied');
    expect(cs2After?.status).toBe('reverted');
    expect(cs3After?.status).toBe('reverted');

    // Replaying the returned inverse actions against the post-change-3 state
    // converges to exactly "change 1 applied, changes 2/3 undone" — the
    // literal DATABASE_SCHEMA.md §6.2 acceptance property, not just status flags.
    const finalValues = replayInverse(['new0', 'new1', 'new2'], result.inverseActions);
    expect(finalValues).toEqual(['new0', 'orig1', 'orig2']);
  });

  it('#29 — fails closed (no writes at all) when a change set in the chain has an irreversible action type', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    const cs1 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'change 1',
      context: buildContext([['orig0']]),
      actions: [setCell(0, 'new0')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs1.changeSetId);

    const checkpoint = await checkpointService.createManual({ workbookId: 'wb-1', conversationId: 'conv-1' });

    // Change 2 uses UPDATE_CHART — reversibility-catalog.ts marks this irreversible (TASKS.md #18):
    // nothing captures a chart's prior configuration to restore. (CREATE_CHART itself became
    // reversible at TASKS.md #15 and no longer fits this fixture's purpose.)
    const cs2 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-2',
      prompt: 'restyle the chart',
      context: buildContext([['new0']]),
      actions: [
        { type: 'UPDATE_CHART', sheetName: 'Sales', chartId: 'Chart1', chartType: 'Line' } as unknown as Action,
      ],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs2.changeSetId);

    await expect(checkpointService.restore(checkpoint.checkpointId)).rejects.toThrow(
      RestoreVerificationError,
    );

    // Fail closed: neither change set was mutated.
    const cs2After = await changeSetService.getById(cs2.changeSetId);
    expect(cs2After?.status).toBe('applied');
    const checkpointAfter = await checkpointModel.findOne({ checkpointId: checkpoint.checkpointId });
    expect((checkpointAfter as unknown as CheckpointDocument | null)?.status).toBe('active');
  });

  it('#29 — restoring a checkpoint with nothing applied since the anchor is a safe no-op', async () => {
    const changeSetModel = createChangeSetModel();
    const checkpointModel = createCheckpointModel();
    const conversationModel = { findOne: jest.fn(() => ({ lean: () => ({ exec: () => Promise.resolve(null) }) })) };
    const trace = workflowTraceStub();

    const checkpointService = new CheckpointService(
      checkpointModel as never,
      changeSetModel as unknown as Model<ChangeSetDocument>,
      trace,
    );
    const changeSetService = new ChangeSetService(
      changeSetModel as unknown as Model<ChangeSetDocument>,
      conversationModel as never,
      trace,
      checkpointService,
    );

    const cs1 = await changeSetService.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'change 1',
      context: buildContext([['orig0']]),
      actions: [setCell(0, 'new0')],
      workbookId: 'wb-1',
    });
    await changeSetService.markApplied(cs1.changeSetId);

    const checkpoint = await checkpointService.createManual({ workbookId: 'wb-1', conversationId: 'conv-1' });
    const result = await checkpointService.restore(checkpoint.checkpointId);

    expect(result.revertedChangeSetIds).toEqual([]);
    expect(result.inverseActions).toEqual([]);
    expect(result.checkpoint.status).toBe('restored');
  });
});

/**
 * Applies SET_CELL inverse actions (diff.engine.ts's beforeStateToInverseActions emits
 * an `address` A1 string, e.g. "A2" — not row/col) to a single-column value array.
 */
function replayInverse(values: string[], inverseActions: Action[]): string[] {
  const next = [...values];
  for (const action of inverseActions) {
    const a = action as unknown as { type: string; address?: string; value?: unknown };
    if (a.type === 'SET_CELL' && typeof a.address === 'string') {
      const rowNumber = Number.parseInt(a.address.replace(/[A-Z]+/i, ''), 10);
      next[rowNumber - 1] = a.value as string;
    }
  }
  return next;
}
