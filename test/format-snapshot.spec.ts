import {
  assertFormatSnapshotsComplete,
  FORMAT_SNAPSHOT_MAX_CELLS,
  FormatSnapshot,
  formatActionArea,
  formatActionsExceedTotalCap,
  formatSnapshotsToInverseActions,
  gridToRects,
  isFormatActionRestorable,
  sanitizeFormatSnapshots,
} from '../src/audit/format-snapshot';
import { RevertNoOpError } from '../src/audit/errors/revert-noop.error';

/** TASKS.md #400 — formatting is reverted from what the add-in read off Excel just before applying it. */

const headerSnapshot = (): FormatSnapshot => ({
  kind: 'format',
  sheetName: 'Monthly Summary',
  restorable: true,
  row: 2,
  col: 0,
  rowCount: 1,
  colCount: 4,
  palette: [{ bold: false, clearFill: true, fontColor: '#000000' }],
  grid: [[0, 0, 0, 0]],
});

describe('gridToRects', () => {
  it('keeps one rectangle for a uniform block', () => {
    expect(gridToRects([[0, 0], [0, 0], [0, 0]])).toEqual([{ row: 0, col: 0, rowCount: 3, colCount: 2, index: 0 }]);
  });

  it('splits a row where the format changes, and extends each run down while the next row repeats it', () => {
    const rects = gridToRects([
      [0, 0, 1],
      [0, 0, 1],
      [0, 0, 2],
    ]);
    expect(rects).toEqual([
      { row: 0, col: 0, rowCount: 3, colCount: 2, index: 0 },
      { row: 0, col: 2, rowCount: 2, colCount: 1, index: 1 },
      { row: 2, col: 2, rowCount: 1, colCount: 1, index: 2 },
    ]);
  });

  it('covers every cell exactly once, even for a checkerboard', () => {
    const grid = [
      [0, 1, 0],
      [1, 0, 1],
    ];
    const rects = gridToRects(grid);
    const seen = grid.map((line) => line.map(() => 0));
    for (const r of rects) {
      for (let i = r.row; i < r.row + r.rowCount; i += 1) {
        for (let j = r.col; j < r.col + r.colCount; j += 1) {
          expect(grid[i][j]).toBe(r.index);
          seen[i][j] += 1;
        }
      }
    }
    expect(seen.flat().every((n) => n === 1)).toBe(true);
  });
});

describe('formatSnapshotsToInverseActions', () => {
  it('turns a snapshot back into the FORMAT_RANGE that puts the old values on the same cells', () => {
    expect(formatSnapshotsToInverseActions([headerSnapshot()])).toEqual([
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

  it('restores the newest snapshot first, so the oldest state is the one that lands last', () => {
    const first = { ...headerSnapshot(), palette: [{ numberFormat: 'General' }] };
    const second = { ...headerSnapshot(), palette: [{ numberFormat: '0.00' }] };
    const inverse = formatSnapshotsToInverseActions([first, second]);
    expect(inverse.map((a) => (a as { format?: { numberFormat?: string } }).format?.numberFormat)).toEqual(['0.00', 'General']);
  });

  it('offsets rectangles from the range origin', () => {
    const inverse = formatSnapshotsToInverseActions([
      {
        kind: 'format',
        sheetName: 'S',
        restorable: true,
        row: 3,
        col: 1,
        rowCount: 2,
        colCount: 2,
        palette: [{ bold: true }, { bold: false }],
        grid: [[0, 1], [0, 1]],
      },
    ]);
    expect(inverse).toEqual([
      expect.objectContaining({ row: 3, col: 1, rowCount: 2, colCount: 1, format: { bold: true } }),
      expect.objectContaining({ row: 3, col: 2, rowCount: 2, colCount: 1, format: { bold: false } }),
    ]);
  });

  it('puts column widths back, joining neighbours that share a width', () => {
    const inverse = formatSnapshotsToInverseActions([
      {
        kind: 'columns',
        sheetName: 'S',
        restorable: true,
        widths: [
          { col: 0, width: 64 },
          { col: 1, width: 64 },
          { col: 2, width: 90 },
          { col: 4, width: 64 },
        ],
      },
    ]);
    expect(inverse).toEqual([
      { type: 'SET_COLUMN_WIDTH', sheetName: 'S', col: 0, colCount: 2, width: 64 },
      { type: 'SET_COLUMN_WIDTH', sheetName: 'S', col: 2, colCount: 1, width: 90 },
      { type: 'SET_COLUMN_WIDTH', sheetName: 'S', col: 4, colCount: 1, width: 64 },
    ]);
  });

  it('skips a snapshot that could not be restored', () => {
    expect(formatSnapshotsToInverseActions([{ kind: 'format', sheetName: 'S', restorable: false, reason: 'x' }])).toEqual([]);
  });
});

describe('isFormatActionRestorable', () => {
  const range = (extra: Record<string, unknown> = {}) => ({
    type: 'FORMAT_RANGE',
    sheetName: 'S',
    row: 2,
    col: 0,
    rowCount: 1,
    colCount: 4,
    format: { bold: true },
    ...extra,
  });

  it('accepts an ordinary formatting action', () => {
    expect(isFormatActionRestorable(range())).toBe(true);
  });

  it('refuses borders, which FORMAT_RANGE cannot put back edge by edge', () => {
    expect(isFormatActionRestorable(range({ format: { borders: 'all' } }))).toBe(false);
  });

  it('refuses a range too large to snapshot, and one whose size cannot be told', () => {
    expect(isFormatActionRestorable(range({ rowCount: 1_000_000, colCount: 16 }))).toBe(false);
    expect(isFormatActionRestorable(range({ rowCount: 'many' }))).toBe(false);
    expect(isFormatActionRestorable({ type: 'FORMAT_RANGE', sheetName: 'S', range: 'A:A', format: { bold: true } })).toBe(false);
  });

  it('reads the size from an A1 range when no indices are given', () => {
    expect(formatActionArea({ type: 'FORMAT_RANGE', range: 'A3:D3' })).toBe(4);
    expect(formatActionArea({ type: 'FORMAT_RANGE', range: 'B4:D11' })).toBe(24);
    expect(formatActionArea({ type: 'FORMAT_RANGE', range: '3:3' })).toBeNull();
  });

  it('accepts the biggest allowed range and refuses one cell more', () => {
    expect(isFormatActionRestorable(range({ rowCount: FORMAT_SNAPSHOT_MAX_CELLS, colCount: 1 }))).toBe(true);
    expect(isFormatActionRestorable(range({ rowCount: FORMAT_SNAPSHOT_MAX_CELLS + 1, colCount: 1 }))).toBe(false);
  });

  it('accepts autofit of the sheet or a few columns, refuses an absurd number of them', () => {
    expect(isFormatActionRestorable({ type: 'AUTOFIT_COLUMNS', sheetName: 'S' })).toBe(true);
    expect(isFormatActionRestorable({ type: 'AUTOFIT_COLUMNS', sheetName: 'S', columns: ['A', 'B'] })).toBe(true);
    expect(isFormatActionRestorable({ type: 'AUTOFIT_COLUMNS', sheetName: 'S', columns: Array.from({ length: 300 }, () => 'A') })).toBe(false);
  });

  it('notices when the formatting of one change set adds up to more than a snapshot may hold', () => {
    const big = range({ rowCount: 20_000, colCount: 1 });
    expect(formatActionsExceedTotalCap([big, big, big])).toBe(false);
    expect(formatActionsExceedTotalCap([big, big, big, big])).toBe(true);
  });
});

describe('assertFormatSnapshotsComplete', () => {
  const actions = [
    { type: 'FORMAT_RANGE', sheetName: 'S', row: 2, col: 0, rowCount: 1, colCount: 4, format: { bold: true } },
    { type: 'AUTOFIT_COLUMNS', sheetName: 'S' },
  ];
  const columns: FormatSnapshot = { kind: 'columns', sheetName: 'S', restorable: true, widths: [{ col: 0, width: 64 }] };

  it('passes when every restorable formatting action has a restorable snapshot', () => {
    expect(() => assertFormatSnapshotsComplete('cs', actions, [headerSnapshot(), columns])).not.toThrow();
  });

  it('refuses when a snapshot never arrived, as a revert error the controller turns into a 422', () => {
    expect(() => assertFormatSnapshotsComplete('cs', actions, [headerSnapshot()])).toThrow(RevertNoOpError);
    expect(() => assertFormatSnapshotsComplete('cs', actions, [])).toThrow(RevertNoOpError);
  });

  it('refuses when any snapshot could not be restored', () => {
    const bad: FormatSnapshot = { kind: 'columns', sheetName: 'S', restorable: false, reason: 'too many columns' };
    expect(() => assertFormatSnapshotsComplete('cs', actions, [headerSnapshot(), bad])).toThrow(RevertNoOpError);
  });

  it('asks for nothing when the change set has no formatting it could restore', () => {
    expect(() => assertFormatSnapshotsComplete('cs', [{ type: 'SET_CELL' }], [])).not.toThrow();
    const borders = [{ type: 'FORMAT_RANGE', row: 0, col: 0, rowCount: 1, colCount: 1, format: { borders: 'all' } }];
    expect(() => assertFormatSnapshotsComplete('cs', borders, [])).not.toThrow();
  });
});

describe('sanitizeFormatSnapshots — the apply call body comes from the add-in', () => {
  it('keeps a well-formed snapshot as it is', () => {
    expect(sanitizeFormatSnapshots([headerSnapshot()])).toEqual([headerSnapshot()]);
  });

  it('returns nothing for something that is not a list', () => {
    expect(sanitizeFormatSnapshots(undefined)).toEqual([]);
    expect(sanitizeFormatSnapshots({ kind: 'format' })).toEqual([]);
  });

  it.each([
    ['a grid whose size is not the range size', { ...headerSnapshot(), grid: [[0, 0]] }],
    ['a grid pointing past the palette', { ...headerSnapshot(), grid: [[0, 0, 0, 5]] }],
    ['a format key it does not know', { ...headerSnapshot(), palette: [{ bold: true, evil: 'x' }] }],
    ['a value of the wrong type', { ...headerSnapshot(), palette: [{ bold: 'yes' }] }],
    ['an alignment it cannot write', { ...headerSnapshot(), palette: [{ horizontalAlignment: 'justify' }] }],
    ['a range larger than a snapshot may be', { ...headerSnapshot(), rowCount: 1000, colCount: 1000, grid: [] }],
    ['a negative origin', { ...headerSnapshot(), row: -1 }],
  ])('turns %s into a refusal, not data', (_label, snapshot) => {
    const [clean] = sanitizeFormatSnapshots([snapshot]);
    expect(clean.restorable).toBe(false);
    expect(clean.palette).toBeUndefined();
    expect(clean.grid).toBeUndefined();
  });

  it('turns a malformed column list into a refusal', () => {
    const [clean] = sanitizeFormatSnapshots([{ kind: 'columns', sheetName: 'S', restorable: true, widths: [{ col: 0, width: -3 }] }]);
    expect(clean.restorable).toBe(false);
  });

  it('keeps a refusal the add-in reported, with its reason', () => {
    const [clean] = sanitizeFormatSnapshots([{ kind: 'format', sheetName: 'S', restorable: false, reason: 'borders' }]);
    expect(clean).toEqual({ kind: 'format', sheetName: 'S', restorable: false, reason: 'borders' });
  });
});
