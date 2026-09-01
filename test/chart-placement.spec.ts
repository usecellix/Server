import {
  applyChartPlacementPass,
  columnIndexToLetter,
  columnLetterToIndex,
  lastColumnOfRange,
  parseCellRef,
} from '../src/excel-ai/utils/chart-placement.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #133 / #163 — a chart must not sit on the table it plots.
 *
 * The failure is from a live run: source `A4:D16`, chart at `D4:K18`, so the
 * chart covered column D — "Pending Amount", a column of its own source. The
 * planner prescribes that anchor literally and the Executor transcribes it
 * faithfully, so the fix is deterministic rather than only a prompt change: it
 * fails silently (no error, a column is just hidden), which is exactly the
 * class of thing a model must not be trusted to remember.
 */
describe('chart placement helpers', () => {
  it('round-trips column letters past Z', () => {
    expect(columnLetterToIndex('A')).toBe(0);
    expect(columnLetterToIndex('D')).toBe(3);
    expect(columnLetterToIndex('AA')).toBe(26);
    expect(columnIndexToLetter(0)).toBe('A');
    expect(columnIndexToLetter(26)).toBe('AA');
  });

  it('parses cell refs, absolute or not', () => {
    expect(parseCellRef('D4')).toEqual({ col: 3, row: 3 });
    expect(parseCellRef('$F$10')).toEqual({ col: 5, row: 9 });
    expect(parseCellRef('nonsense')).toBeNull();
  });

  it('finds a range last column, with or without a sheet prefix', () => {
    expect(lastColumnOfRange('A4:D16')).toBe(3);
    expect(lastColumnOfRange('Main!A4:D16')).toBe(3);
    expect(lastColumnOfRange('B2')).toBe(1);
  });
});

describe('applyChartPlacementPass', () => {
  function chart(overrides: Partial<SheetActionPayload> = {}): SheetActionPayload {
    return {
      type: 'CREATE_CHART',
      sheetName: 'Main',
      sourceSheetName: 'Main',
      sourceRange: 'A4:D16',
      chartType: 'ColumnClustered',
      title: 'Monthly Totals',
      startCell: 'D4',
      endCell: 'K18',
      ...overrides,
    } as SheetActionPayload;
  }

  it('moves a chart off its own source columns — the exact live failure', () => {
    const [moved] = applyChartPlacementPass([chart()]);
    // Source ends at D (index 3); first free column with a gutter is F.
    expect(moved.startCell).toBe('F4');
  });

  it('preserves the chart width when it moves', () => {
    const [moved] = applyChartPlacementPass([chart()]);
    // D->F is +2 columns, so K->M.
    expect(moved.endCell).toBe('M18');
  });

  it('keeps the chart on its original row', () => {
    const [moved] = applyChartPlacementPass([chart({ startCell: 'B10', endCell: 'I24' } as never)]);
    expect(moved.startCell).toMatch(/10$/);
    expect(moved.endCell).toMatch(/24$/);
  });

  it('leaves a chart that is already clear of its source untouched', () => {
    const input = [chart({ startCell: 'G4', endCell: 'N18' } as never)];
    expect(applyChartPlacementPass(input)).toBe(input);
  });

  it('handles a narrow source correctly — the prompt example stays valid', () => {
    // source A4:B9 ends at B (1); first free with gutter is D — so D4 is fine.
    const input = [chart({ sourceRange: 'A4:B9', startCell: 'D4' } as never)];
    expect(applyChartPlacementPass(input)).toBe(input);
  });

  it('never moves a chart whose source is on a DIFFERENT sheet', () => {
    const input = [
      chart({ sheetName: 'Dashboard', sourceSheetName: 'Data', startCell: 'A1' } as never),
    ];
    expect(applyChartPlacementPass(input)).toBe(input);
  });

  it('leaves every non-chart action alone', () => {
    const others: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'x' },
      { type: 'FREEZE_PANES', sheetName: 'Main', freezeRows: 1 },
    ];
    const out = applyChartPlacementPass([...others, chart()]);
    expect(out.slice(0, 2)).toEqual(others);
    expect(out).toHaveLength(3);
  });

  it('is a no-op for a batch with no charts', () => {
    const input: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'x' },
    ];
    expect(applyChartPlacementPass(input)).toBe(input);
  });

  it('tolerates a malformed anchor rather than throwing', () => {
    const input = [chart({ startCell: '', sourceRange: '' } as never)];
    expect(() => applyChartPlacementPass(input)).not.toThrow();
  });
});
