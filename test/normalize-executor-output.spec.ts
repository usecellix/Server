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

  it('normalizes a raw CONDITIONAL_FORMAT action, preserving its rule (TASKS.md #33)', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Purchase Register',
            range: "'Purchase Register'!J2:J51",
            rule: {
              kind: 'cellValue',
              operator: 'greaterThan',
              value: 1000,
              format: { fillColor: '#FFC7CE' },
            },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51', // sheet prefix stripped
        rule: {
          kind: 'cellValue',
          operator: 'greaterThan',
          value: 1000,
          format: { fillColor: '#FFC7CE' },
        },
      }),
    ]);
  });

  it("normalizes a formula-kind CONDITIONAL_FORMAT (TASKS.md #35 — VISION.md's own example: revenue dropped >10%)", () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Regional Revenue',
            range: 'A2:D9',
            rule: {
              kind: 'formula',
              formula: '=$C2<$B2*0.9',
              format: { fillColor: '#FFC7CE' },
            },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: { kind: 'formula', formula: '=$C2<$B2*0.9', format: { fillColor: '#FFC7CE' } },
      }),
    ]);
  });

  it('drops a formula-kind CONDITIONAL_FORMAT with an empty formula (fails closed)', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Regional Revenue',
            range: 'A2:D9',
            rule: { kind: 'formula', formula: '   ', format: { fillColor: '#FFC7CE' } },
          },
        ],
      },
      subtask,
    );
    expect(result.actions).toEqual([]);
    expect((result.droppedActions ?? []).length).toBeGreaterThan(0);
  });

  it('normalizes a topBottom-kind CONDITIONAL_FORMAT (TASKS.md #36 — "highlight the top 5 suppliers by total")', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Suppliers',
            range: 'C2:C40',
            rule: {
              kind: 'topBottom',
              side: 'top',
              rank: 5,
              format: { fillColor: '#C6EFCE' },
            },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 5, format: { fillColor: '#C6EFCE' } },
      }),
    ]);
  });

  it('normalizes a percent-based bottom topBottom rule, preserving isPercent', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Scores',
            range: 'B2:B50',
            rule: {
              kind: 'topBottom',
              side: 'bottom',
              rank: 10,
              isPercent: true,
              format: { fillColor: '#FFC7CE' },
            },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        rule: {
          kind: 'topBottom',
          side: 'bottom',
          rank: 10,
          isPercent: true,
          format: { fillColor: '#FFC7CE' },
        },
      }),
    ]);
  });

  it('drops a topBottom CONDITIONAL_FORMAT with an invalid side or non-positive rank (fails closed)', () => {
    const badSide = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Suppliers',
            range: 'C2:C40',
            rule: { kind: 'topBottom', side: 'middle', rank: 5, format: { fillColor: '#C6EFCE' } },
          },
        ],
      },
      subtask,
    );
    expect(badSide.actions).toEqual([]);
    expect((badSide.droppedActions ?? []).length).toBeGreaterThan(0);

    const badRank = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Suppliers',
            range: 'C2:C40',
            rule: { kind: 'topBottom', side: 'top', rank: 0, format: { fillColor: '#C6EFCE' } },
          },
        ],
      },
      subtask,
    );
    expect(badRank.actions).toEqual([]);
    expect((badRank.droppedActions ?? []).length).toBeGreaterThan(0);
  });

  it('normalizes a colorScale-kind CONDITIONAL_FORMAT (TASKS.md #37 — "color-scale the Total Amount column")', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Scores',
            range: 'B2:B50',
            rule: { kind: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
      }),
    ]);
  });

  it('normalizes a 2-color colorScale rule', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Scores',
            range: 'B2:B50',
            rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({
        rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      }),
    ]);
  });

  it('drops a colorScale CONDITIONAL_FORMAT with the wrong number of colors or a non-string entry (fails closed)', () => {
    const wrongCount = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Scores',
            range: 'B2:B50',
            rule: { kind: 'colorScale', colors: ['#F8696B'] },
          },
        ],
      },
      subtask,
    );
    expect(wrongCount.actions).toEqual([]);
    expect((wrongCount.droppedActions ?? []).length).toBeGreaterThan(0);

    const badEntry = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Scores',
            range: 'B2:B50',
            rule: { kind: 'colorScale', colors: ['#F8696B', 123] },
          },
        ],
      },
      subtask,
    );
    expect(badEntry.actions).toEqual([]);
    expect((badEntry.droppedActions ?? []).length).toBeGreaterThan(0);
  });

  it('preserves existingRuleId on a CONDITIONAL_FORMAT action, so it modifies rather than duplicates (TASKS.md #38)', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Purchase Register',
            range: 'J2:J51',
            existingRuleId: 'cf-abc123',
            rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
          },
        ],
      },
      subtask,
    );

    expect(result.actions).toEqual([
      expect.objectContaining({ existingRuleId: 'cf-abc123' }),
    ]);
  });

  it('drops a blank/whitespace-only existingRuleId rather than passing it through', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'CONDITIONAL_FORMAT',
            sheetName: 'Purchase Register',
            range: 'J2:J51',
            existingRuleId: '   ',
            rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]).not.toHaveProperty('existingRuleId');
  });

  it('drops a CONDITIONAL_FORMAT action missing its rule (fails closed, not silently accepted)', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [{ type: 'CONDITIONAL_FORMAT', sheetName: 'Purchase Register', range: 'J2:J51' }],
      },
      subtask,
    );

    expect(result.actions).toEqual([]);
    expect((result.droppedActions ?? []).length).toBeGreaterThan(0);
  });

  it('no longer collapses a real CONDITIONAL_FORMAT type into FORMAT_MATCHING_ROWS', () => {
    const result = normalizeExecutorOutput(
      {
        actions: [
          {
            type: 'conditional_format',
            sheetName: 'Purchase Register',
            range: 'J2:J51',
            rule: { kind: 'cellValue', operator: 'lessThan', value: 100, format: { fillColor: '#FFC7CE' } },
          },
        ],
      },
      subtask,
    );

    expect(result.actions[0]?.type).toBe('CONDITIONAL_FORMAT');
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

    // TASKS.md #322 — live: the model emitted `operations: [28]` (the COUNT of
    // the writes it meant) for Main's Monthly Totals. It passed as non-empty,
    // the subtask "completed", and rows 4–10 were never written.
    it.each([
      ['a bare count', [28]],
      ['strings instead of objects', ['A4', 'B4']],
      ['an op naming no cell', [{ value: 'Month' }]],
      ['an op writing nothing', [{ address: 'A4' }]],
      ['one good op mixed with a bad one', [{ address: 'A4', value: 'Month' }, 14]],
    ])('drops a BATCH_SET whose operations are not cell writes: %s', (_label, operations) => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'BATCH_SET', sheetName: 'Main', operations }] },
        subtask,
      );

      expect(result.actions).toHaveLength(0);
      expect(result.droppedActions).toEqual([
        { rawType: 'BATCH_SET', reason: 'missing-required-fields' },
      ]);
    });

    it('keeps a BATCH_SET whose operations use row/col instead of an address', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [
            { type: 'BATCH_SET', sheetName: 'Main', operations: [{ row: 3, col: 0, value: 'Month' }] },
          ],
        },
        subtask,
      );
      expect(result.actions).toHaveLength(1);
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

  /**
   * TASKS.md #265 — `Range.format.columnWidth` is in POINTS (default column
   * ~48pt), but the Executor prompt's own examples disagreed on the unit (130
   * vs. 20), and the model reached for small "character count"-looking
   * numbers styling a live dashboard: widths of 12-22 points made every
   * header ("Guest Name", "Rate Per Night"...) clip down to 1-2 characters,
   * reading as a blank/broken sheet.
   */
  describe('SET_COLUMN_WIDTH width clamp (TASKS.md #265)', () => {
    /**
     * TASKS.md #273 — this used to assert a flat floor (14 -> 40). A later
     * live run emitted a sub-40 value for ALL THIRTEEN columns, so every one
     * floored to exactly 40 and the sheet came out uniformly cramped: the
     * floor destroyed the relative sizing the model had got right. The value
     * is now read as a character count and converted, which both widens it
     * and keeps a wide column wide.
     */
    it('converts a character-count width into points instead of flattening it', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 2, width: 14 }],
        },
        subtask,
      );
      expect(result.actions).toHaveLength(1);
      // 14 chars -> 14*7+5 = 103px -> 77pt.
      expect((result.actions[0] as { width: number }).width).toBe(77);
    });

    it("round-trips Excel's own default column width (8.43 chars is 48pt)", () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 0, width: 8.43 }],
        },
        subtask,
      );
      expect((result.actions[0] as { width: number }).width).toBe(48);
    });

    it('preserves RELATIVE sizing — a wide name column stays wider than a narrow code column', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [
            { type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 0, width: 8 },
            { type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 2, width: 22 },
          ],
        },
        subtask,
      );
      const [unitNo, guestName] = result.actions as unknown as Array<{ width: number }>;
      expect(guestName.width).toBeGreaterThan(unitNo.width);
      // The exact failure being fixed: these must NOT both be 40.
      expect(unitNo.width).not.toBe(guestName.width);
    });

    it('still floors a degenerate tiny count that converts below readability', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 2, width: 2 }],
        },
        subtask,
      );
      expect((result.actions[0] as { width: number }).width).toBe(40);
    });

    it('leaves an already-reasonable width untouched', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 2, width: 130 }],
        },
        subtask,
      );
      expect((result.actions[0] as { width: number }).width).toBe(130);
    });

    it('does not touch width on an unrelated action type', () => {
      // width is a generic optional scalar copied onto every action type —
      // the clamp must only fire for SET_COLUMN_WIDTH.
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [
            { type: 'BATCH_SET', sheetName: 'January', operations: [{ address: 'A1', value: 1 }], width: 5 },
          ],
        },
        subtask,
      );
      expect((result.actions[0] as { width?: number }).width).toBe(5);
    });

    it('leaves a missing or non-numeric width for sanitizeAction to reject, rather than inventing one', () => {
      const result = normalizeExecutorOutput(
        {
          subtaskId: 's1',
          actions: [{ type: 'SET_COLUMN_WIDTH', sheetName: 'January', col: 2 }],
        },
        subtask,
      );
      expect((result.actions[0] as { width?: number }).width).toBeUndefined();
    });
  });

  /**
   * Live incident (TASKS.md #279): frontend.log showed 5 ADD_SHEET actions in
   * one changeset throwing "RichApi.Error: The argument is invalid or missing
   * or has an incorrect format." — the client handler
   * (sheet.handler.ts:handleAddSheet) reads only `action.name` for the new
   * sheet's name, but the model, following this normalizer's own generic
   * "sheetName" convention used by every other action type, sometimes emits
   * ADD_SHEET with `sheetName` and no `name` at all.
   */
  describe('ADD_SHEET/CREATE_SHEET name fallback (TASKS.md #279)', () => {
    it('falls back to sheetName when the model omits name entirely', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'ADD_SHEET', sheetName: 'May' }] },
        subtask,
      );
      expect((result.actions[0] as { name?: string }).name).toBe('May');
    });

    it('does the same for CREATE_SHEET', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'CREATE_SHEET', sheetName: 'May' }] },
        subtask,
      );
      expect((result.actions[0] as { name?: string }).name).toBe('May');
    });

    it('leaves an explicit name untouched — never overrides a correct emission', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'ADD_SHEET', sheetName: 'May', name: 'June' }] },
        subtask,
      );
      expect((result.actions[0] as { name?: string }).name).toBe('June');
    });

    it('never applies the fallback to other action types (sheetName means something else there)', () => {
      const result = normalizeExecutorOutput(
        { subtaskId: 's1', actions: [{ type: 'SET_CELL', sheetName: 'May', address: 'A1', value: 1 }] },
        subtask,
      );
      expect((result.actions[0] as { name?: string }).name).toBeUndefined();
    });
  });
});
