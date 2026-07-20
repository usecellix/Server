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
});
