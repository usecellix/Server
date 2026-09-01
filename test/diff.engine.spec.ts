import { WorkbookContext } from '../src/agents/types/agent.types';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { virtualApply } from '../src/virtual/virtualApply';
import {
  beforeStateToInverseActions,
  computeUnintendedChanges,
  detectIntroducedFormulaErrors,
  diffShadowsFully,
  generateDiff,
  shadowFromBeforeState,
  snapshotBeforeState,
} from '../src/audit/diff.engine';
import { CellChange } from '../src/audit/types/change-set.types';
import { Action } from '../src/agents/types/agent.types';

const baseContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C2',
      rowCount: 2,
      columnCount: 3,
      values: [
        ['Item', 'Qty', 'Price'],
        ['Apple', 10, 1.5],
      ],
      formulas: [['', '', ''], ['', '', '']],
      numberFormats: [['General', 'General', 'General']],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('diff.engine', () => {
  it('snapshots before state from shadow workbook', () => {
    const shadow = buildShadowWorkbook(baseContext);
    const state = snapshotBeforeState(shadow);
    expect(state['Sheet1!A1']).toEqual({ value: 'Item', formula: '', format: 'General' });
    expect(state['Sheet1!C2']).toEqual({ value: 1.5, formula: '', format: 'General' });
  });

  it('generates cell-level diff after virtual apply', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    const actions = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 },
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2*1.1' },
    ];
    const after = virtualApply(before, actions as never);
    const changes = generateDiff(before, after);

    expect(changes.length).toBeGreaterThanOrEqual(2);
    const qtyChange = changes.find((c: CellChange) => c.cell === 'B2');
    expect(qtyChange).toMatchObject({ sheet: 'Sheet1', before: 10, after: 12 });

    const inverse = beforeStateToInverseActions(beforeState, changes);
    expect(inverse.some((a: Action) => a.type === 'SET_CELL' && (a as Action & { address?: string }).address === 'B2')).toBe(true);
  });

  it('produces inverse actions that restore before values', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    const after = virtualApply(before, [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'Product' },
    ] as never);
    const changes = generateDiff(before, after);
    const inverse = beforeStateToInverseActions(beforeState, changes);

    const restored = virtualApply(after, inverse as never);
    expect(restored.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
  });

  it('restores the captured number format alongside the value on revert (TASKS.md #10)', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    expect(beforeState['Sheet1!C2']).toMatchObject({ format: 'General' });

    // Forward change touches both value and format — e.g. an agent turning a
    // plain number into a percentage.
    const after = virtualApply(before, [
      {
        type: 'SET_CELL',
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        value: 0.15,
        format: { numberFormat: '0.00%' },
      },
    ] as never);
    expect(after.sheets.get('Sheet1')?.cells.get('C2')?.numberFormat).toBe('0.00%');

    const changes = generateDiff(before, after);
    const inverse = beforeStateToInverseActions(beforeState, changes);
    const inverseForC2 = inverse.find(
      (a) => (a as Action & { address?: string }).address === 'C2',
    ) as (Action & { format?: { numberFormat?: string } }) | undefined;
    expect(inverseForC2?.format?.numberFormat).toBe('General');

    const restored = virtualApply(after, inverse as never);
    const restoredCell = restored.sheets.get('Sheet1')?.cells.get('C2');
    expect(restoredCell?.value).toBe(1.5);
    // Without the #10 fix, this stays '0.00%' — the inverse action never carried
    // a format field, so virtualSetCell preserved whatever format the forward
    // change left behind instead of restoring the captured before-state.
    expect(restoredCell?.numberFormat).toBe('General');
  });

  it('restores bold/italic/fontColor/fillColor alongside value and number format on revert (TASKS.md #64)', () => {
    const before = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(before);
    // C2 starts with no captured bold/italic/fontColor/fillColor (baseContext has no
    // `formats` field) — the real "before" state a plain, unformatted cell would have.
    expect(beforeState['Sheet1!C2']).toMatchObject({
      format: 'General',
      bold: undefined,
      italic: undefined,
      fontColor: undefined,
      fillColor: undefined,
    });

    // Forward change: bold + fill + a non-default number format in one request —
    // the exact scenario TASKS.md #64's own acceptance test describes.
    const after = virtualApply(before, [
      {
        type: 'SET_CELL',
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        value: 0.15,
        format: { numberFormat: '0.00%', bold: true, fillColor: '#FFFF00' },
      },
    ] as never);
    expect(after.sheets.get('Sheet1')?.cells.get('C2')).toMatchObject({
      numberFormat: '0.00%',
      bold: true,
      fillColor: '#FFFF00',
    });

    const changes = generateDiff(before, after);
    const inverse = beforeStateToInverseActions(beforeState, changes);
    const inverseForC2 = inverse.find(
      (a) => (a as Action & { address?: string }).address === 'C2',
    ) as (Action & { format?: Record<string, unknown> }) | undefined;
    // The inverse must not force bold/fill to false/empty — nothing was ever
    // captured for them (the cell had no formatting to begin with), so they must
    // be entirely absent from the restore action, not present-and-falsy. A
    // present-and-falsy key would still be "wrong" in a different way: it assumes
    // "not captured" and "known to be unformatted" are the same thing, which they
    // aren't (see CellFormatCell's own doc comment) — but the two are indistinguishable
    // in *this* specific test's fixture (baseContext genuinely has neither formats nor
    // any bold cell), so the sharper assertion here is simply "absent".
    expect(inverseForC2?.format).not.toHaveProperty('bold');
    expect(inverseForC2?.format).not.toHaveProperty('fillColor');
    expect(inverseForC2?.format?.numberFormat).toBe('General');

    const restored = virtualApply(after, inverse as never);
    const restoredCell = restored.sheets.get('Sheet1')?.cells.get('C2');
    expect(restoredCell?.value).toBe(1.5);
    expect(restoredCell?.numberFormat).toBe('General');
    // Known, documented limitation (not a bug): when the source SheetContext never
    // captured bold/fillColor at all (no `formats` field — an older add-in build,
    // or the minimal-context fallback), revert has no prior value to restore *to*,
    // so it deliberately leaves whatever the forward change set rather than forcing
    // a guessed default — forcing `bold: false` could just as easily clobber a real
    // prior `true` the context simply never captured (see the next test for the case
    // where the prior value *is* known, which restores correctly).
    expect(restoredCell?.bold).toBe(true);
    expect(restoredCell?.fillColor).toBe('#FFFF00');
  });

  it('restores a cell that was ALREADY bold/filled back to its real prior formatting, not just clears it (TASKS.md #64)', () => {
    // Header row (row 0) is bold+filled to begin with — via `formats`, the new
    // SheetContext field TASKS.md #64 introduced (column-broadcast, same
    // granularity as numberFormats).
    const formattedContext: WorkbookContext = {
      ...baseContext,
      sheets: [
        {
          ...baseContext.sheets[0]!,
          formats: [
            [
              { bold: true, fillColor: '#D9D9D9' },
              { bold: true, fillColor: '#D9D9D9' },
              { bold: true, fillColor: '#D9D9D9' },
            ],
            [{}, {}, {}],
          ],
        },
      ],
    };

    const before = buildShadowWorkbook(formattedContext);
    const beforeState = snapshotBeforeState(before);
    expect(beforeState['Sheet1!A1']).toMatchObject({ bold: true, fillColor: '#D9D9D9' });

    // Forward change strips the header's bold/fill and renames it — e.g. an
    // agent restyling the header, or a plain FILL_DOWN-style overwrite.
    const after = virtualApply(before, [
      {
        type: 'SET_CELL',
        sheetName: 'Sheet1',
        row: 0,
        col: 0,
        value: 'Product',
        format: { bold: false, fillColor: '#FFFFFF' },
      },
    ] as never);
    expect(after.sheets.get('Sheet1')?.cells.get('A1')).toMatchObject({
      value: 'Product',
      bold: false,
      fillColor: '#FFFFFF',
    });

    const changes = generateDiff(before, after);
    const inverse = beforeStateToInverseActions(beforeState, changes);
    const restored = virtualApply(after, inverse as never);
    const restoredCell = restored.sheets.get('Sheet1')?.cells.get('A1');

    expect(restoredCell?.value).toBe('Item');
    // The real point of this test: restored to `true`, not merely "not false" —
    // proves the captured original value round-trips, not just that *some*
    // format key got attached to the inverse action.
    expect(restoredCell?.bold).toBe(true);
    expect(restoredCell?.fillColor).toBe('#D9D9D9');
  });

  it('shadowFromBeforeState rebuilds a shadow that round-trips through snapshotBeforeState (TASKS.md #19)', () => {
    const shadow = buildShadowWorkbook(baseContext);
    const beforeState = snapshotBeforeState(shadow);

    const rebuilt = shadowFromBeforeState(beforeState);

    expect(rebuilt.sheets.get('Sheet1')?.cells.get('A1')?.value).toBe('Item');
    expect(rebuilt.sheets.get('Sheet1')?.cells.get('C2')?.value).toBe(1.5);
    expect(diffShadowsFully(shadow, rebuilt)).toEqual([]);
  });

  it('diffShadowsFully catches a mismatch that changedCells-based generateDiff would miss (TASKS.md #19)', () => {
    const shadow = buildShadowWorkbook(baseContext);
    // One real, tracked change (marks changedCells)...
    const afterRealChange = virtualApply(shadow, [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'Product' },
    ] as never);
    // ...plus a second, untracked mutation applied directly (not via virtualApply, so
    // changedCells is never marked for it) — this is exactly the "an inverse action
    // forgot a cell" failure mode revert self-verification needs to catch.
    const sheet = afterRealChange.sheets.get('Sheet1')!;
    sheet.cells.set('B2', { value: 999, formula: '', numberFormat: 'General' });

    // The fast path (changedCells is non-empty from the tracked change) misses the
    // untracked mutation entirely.
    const fastDiff = generateDiff(shadow, afterRealChange);
    expect(fastDiff.some((c) => c.cell === 'B2')).toBe(false);

    // The full comparison correctly catches both.
    const blocking = diffShadowsFully(shadow, afterRealChange);
    expect(blocking.some((c) => c.cell === 'A1')).toBe(true);
    const b2Blocker = blocking.find((c) => c.cell === 'B2');
    expect(b2Blocker).toMatchObject({ sheet: 'Sheet1', before: 10, after: 999 });
  });

  describe('computeUnintendedChanges (TASKS.md #48 — PRD A5)', () => {
    it('flags a cell change on a sheet no action in the batch declared touching', () => {
      const changes: CellChange[] = [
        { cell: 'B2', sheet: 'Sheet1', before: 10, after: 12, isHardcoded: true },
        { cell: 'A1', sheet: 'Sheet2', before: null, after: 'surprise', isHardcoded: true },
      ];
      const actions = [
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 },
      ] as Action[];

      const unintended = computeUnintendedChanges(changes, actions);
      expect(unintended).toHaveLength(1);
      expect(unintended[0]).toMatchObject({ sheet: 'Sheet2', cell: 'A1' });
    });

    it('counts zero when every change fell on a declared sheet', () => {
      const changes: CellChange[] = [
        { cell: 'B2', sheet: 'Sheet1', before: 10, after: 12, isHardcoded: true },
        { cell: 'D4', sheet: 'Sheet1', before: 1, after: 2, isHardcoded: true },
      ];
      const actions = [
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 1, col: 1, value: 12 },
        { type: 'SET_CELL', sheetName: 'Sheet1', row: 3, col: 3, value: 2 },
      ] as Action[];

      expect(computeUnintendedChanges(changes, actions)).toEqual([]);
    });

    it('fails open (reports nothing) when no action in the batch declares a sheetName', () => {
      const changes: CellChange[] = [
        { cell: 'B2', sheet: 'Sheet1', before: 10, after: 12, isHardcoded: true },
      ];
      const actions = [{ type: 'CHECKPOINT', message: 'manual checkpoint' }] as Action[];

      expect(computeUnintendedChanges(changes, actions)).toEqual([]);
    });
  });

  describe('detectIntroducedFormulaErrors (TASKS.md #49 — PRD A6)', () => {
    it('detects a newly introduced #REF! error and attributes it to the right cell', () => {
      const changes: CellChange[] = [
        { cell: 'C3', sheet: 'Sheet1', before: 5, after: '#REF!', isHardcoded: false },
      ];
      const introduced = detectIntroducedFormulaErrors(changes);
      expect(introduced).toEqual([{ cell: 'C3', sheet: 'Sheet1', error: '#REF!' }]);
    });

    it('does not double-count a pre-existing, unrelated error elsewhere in the change set', () => {
      const changes: CellChange[] = [
        // Genuinely new — was a plain value, now errors.
        { cell: 'C3', sheet: 'Sheet1', before: 5, after: '#REF!', isHardcoded: false },
        // Already errored before this change touched it (e.g. re-sorted into place) —
        // the exact same error, not a new occurrence.
        { cell: 'D9', sheet: 'Sheet1', before: '#DIV/0!', after: '#DIV/0!', isHardcoded: false },
      ];
      const introduced = detectIntroducedFormulaErrors(changes);
      expect(introduced).toEqual([{ cell: 'C3', sheet: 'Sheet1', error: '#REF!' }]);
    });

    it('counts zero for an in-scope run with no formula errors', () => {
      const changes: CellChange[] = [
        { cell: 'B2', sheet: 'Sheet1', before: 10, after: 12, isHardcoded: true },
      ];
      expect(detectIntroducedFormulaErrors(changes)).toEqual([]);
    });
  });
});
