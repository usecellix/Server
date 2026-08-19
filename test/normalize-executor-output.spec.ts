import { normalizeExecutorOutput } from '../src/agents/utils/normalize-executor-output.util';
import { SubTask } from '../src/agents/types/agent.types';

describe('normalizeExecutorOutput', () => {
  const subtask: SubTask = {
    id: 's1',
    description: 'Create a table',
    targetSheet: 'Purchase Register',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('canonicalizes a missing or invented subtaskId to the active subtask', () => {
    expect(normalizeExecutorOutput({ subtaskId: '', actions: [] }, subtask).subtaskId).toBe(
      's1',
    );
    expect(
      normalizeExecutorOutput({ subtaskId: 'Subtask: create table', actions: [] }, subtask)
        .subtaskId,
    ).toBe('s1');
  });

  it('canonicalizes legacy CREATE_TABLE name and defaults headers to true', () => {
    const result = normalizeExecutorOutput(
      {
        subtaskId: 's1',
        actions: [
          {
            type: 'CREATE_TABLE',
            sheetName: 'Purchase Register',
            range: 'A1:L51',
            name: 'PurchaseTable',
          },
        ],
        isDone: true,
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        type: 'CREATE_TABLE',
        tableName: 'PurchaseTable',
        hasHeaders: true,
      }),
    ]);
  });

  it('preserves an explicit hasHeaders false value', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CREATE_TABLE',
            range: 'A1:B2',
            tableName: 'RawTable',
            hasHeaders: false,
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).toEqual(
      expect.objectContaining({ tableName: 'RawTable', hasHeaders: false }),
    );
  });

  it('converts FORMAT_RANGE A1 range string into row/col/rowCount/colCount', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'FORMAT_RANGE',
            sheetName: 'X',
            range: 'A1:L1',
            format: { bold: true, fillColor: '#FF0000' },
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).toEqual(
      expect.objectContaining({
        type: 'FORMAT_RANGE',
        sheetName: 'X',
        row: 0,
        col: 0,
        rowCount: 1,
        colCount: 12,
        format: { bold: true, fillColor: '#FF0000' },
      }),
    );
  });

  it('preserves CREATE_CHART source and placement fields', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CREATE_CHART',
            sheetName: 'Dashboard',
            sourceSheetName: 'Purchase Register',
            sourceRange: 'A1:B10',
            chartType: 'Line',
            title: 'Monthly Purchases',
            startCell: 'A8',
            endCell: 'H24',
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).toEqual(
      expect.objectContaining({
        type: 'CREATE_CHART',
        sourceSheetName: 'Purchase Register',
        sourceRange: 'A1:B10',
        chartType: 'Line',
        title: 'Monthly Purchases',
        startCell: 'A8',
        endCell: 'H24',
      }),
    );
  });

  it('normalizes COPY_FILTERED_RANGE fields and defaults', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'COPY_FILTERED_RANGE',
            sourceSheet: 'Purchase Register',
            sourceRange: 'A1:L51',
            destSheet: 'Pending Payments',
            destStartCell: 'A1',
            filter: {
              column: 'Payment Status',
              operator: 'equals',
              value: 'Pending',
            },
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).toEqual(
      expect.objectContaining({
        type: 'COPY_FILTERED_RANGE',
        sourceSheet: 'Purchase Register',
        sourceRange: 'A1:L51',
        destSheet: 'Pending Payments',
        destStartCell: 'A1',
        hasHeaders: true,
        mode: 'copy',
        filter: {
          column: 'Payment Status',
          operator: 'equals',
          value: 'Pending',
        },
      }),
    );
  });

  it('normalizes MOVE_RANGE fields', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'MOVE_RANGE',
            sourceSheet: 'Sheet1',
            sourceRange: 'A1:D10',
            destSheet: 'Archive',
            destStartCell: 'B2',
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).toEqual(
      expect.objectContaining({
        type: 'MOVE_RANGE',
        sourceSheet: 'Sheet1',
        sourceRange: 'A1:D10',
        destSheet: 'Archive',
        destStartCell: 'B2',
      }),
    );
  });

  describe('unusable actions are reported, not silently discarded', () => {
    it('reports an unknown action type instead of dropping it in silence', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [
            { type: 'SET_CELL', address: 'A1', value: 'Total' },
            { type: 'APPLY_PIVOT_MAGIC', range: 'A1:D10' },
          ],
        },
        subtask,
      );

      expect(result.actions).toHaveLength(1);
      expect(result.droppedActions).toEqual([
        { rawType: 'APPLY_PIVOT_MAGIC', reason: 'unknown-type' },
      ]);
    });

    it('reports a non-object action entry', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: ['FREEZE_PANES', null] },
        subtask,
      );

      expect(result.actions).toHaveLength(0);
      expect(result.droppedActions).toEqual([
        { rawType: null, reason: 'not-an-object' },
        { rawType: null, reason: 'not-an-object' },
      ]);
    });

    it('leaves droppedActions empty when every action normalizes', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'SET_CELL', address: 'A1', value: 1 }] },
        subtask,
      );

      expect(result.actions).toHaveLength(1);
      expect(result.droppedActions).toEqual([]);
    });

    // Regression: a BATCH_SET with no operations array reached the frontend
    // undiscarded (every downstream check guards with Array.isArray and just
    // skips its own logic instead of rejecting), and crashed the entire apply —
    // "action.operations is not iterable" — losing every other verified action
    // in the same batch. This is a recognized type with a missing required
    // field, not an unknown type, so it needs its own reason category.
    it('drops a BATCH_SET with no operations array, with reason missing-required-fields', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'BATCH_SET', sheetName: 'Main' }] },
        subtask,
      );

      expect(result.actions).toHaveLength(0);
      expect(result.droppedActions).toEqual([
        { rawType: 'BATCH_SET', reason: 'missing-required-fields' },
      ]);
    });

    it('drops a BATCH_SET with an empty operations array', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'BATCH_SET', sheetName: 'Main', operations: [] }] },
        subtask,
      );

      expect(result.actions).toHaveLength(0);
      expect(result.droppedActions).toEqual([
        { rawType: 'BATCH_SET', reason: 'missing-required-fields' },
      ]);
    });

    it('drops a BATCH_SET whose operations field is the wrong shape (not an array)', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'BATCH_SET', sheetName: 'Main', operations: { address: 'B2' } }],
        },
        subtask,
      );

      expect(result.actions).toHaveLength(0);
      expect(result.droppedActions).toEqual([
        { rawType: 'BATCH_SET', reason: 'missing-required-fields' },
      ]);
    });

    it('keeps a well-formed BATCH_SET with a non-empty operations array', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [
            {
              type: 'BATCH_SET',
              sheetName: 'Main',
              operations: [{ address: 'B2', formula: '=SUM(January!G:G)' }],
            },
          ],
        },
        subtask,
      );

      expect(result.actions).toHaveLength(1);
      expect(result.droppedActions).toEqual([]);
    });
  });
});
