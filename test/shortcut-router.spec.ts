import { routeShortcutAction, hasConditionalShortcutBlocker } from '../src/excel-ai/utils/shortcut-router.util';

describe('routeShortcutAction', () => {
  it('routes freeze top row command', () => {
    expect(routeShortcutAction('freeze top row', 'Sheet1')).toEqual([
      { type: 'FREEZE_PANES', sheetName: 'Sheet1', freezeRows: 1, freezeColumns: 0 },
    ]);
  });

  it('routes freeze the first row command', () => {
    expect(routeShortcutAction('freeze the first row', 'Sheet1')).toEqual([
      { type: 'FREEZE_PANES', sheetName: 'Sheet1', freezeRows: 1, freezeColumns: 0 },
    ]);
  });

  it('routes freeze first N rows', () => {
    expect(routeShortcutAction('freeze first 3 rows', 'Sheet1')).toEqual([
      { type: 'FREEZE_PANES', sheetName: 'Sheet1', freezeRows: 3, freezeColumns: 0 },
    ]);
  });

  it('routes freeze rows 1 through N', () => {
    expect(routeShortcutAction('freeze rows 1 through 3', 'Sheet1')).toEqual([
      { type: 'FREEZE_PANES', sheetName: 'Sheet1', freezeRows: 3, freezeColumns: 0 },
    ]);
  });

  it('routes hide row span command', () => {
    expect(routeShortcutAction('hide rows 10 through 20')).toEqual([
      { type: 'HIDE_ROW', sheetName: undefined, row: 9, rowCount: 11 },
    ]);
  });

  it('routes hide row 7', () => {
    expect(routeShortcutAction('hide row 7', 'Sheet1')).toEqual([
      { type: 'HIDE_ROW', sheetName: 'Sheet1', row: 6, rowCount: 1 },
    ]);
  });

  it('routes unhide column command', () => {
    expect(routeShortcutAction('unhide column D')).toEqual([
      { type: 'UNHIDE_COLUMN', sheetName: undefined, col: 3, colCount: 1 },
    ]);
  });

  it('routes zoom command', () => {
    expect(routeShortcutAction('zoom to 150%')).toEqual([
      { type: 'SET_ZOOM', sheetName: undefined, zoomPercent: 150 },
    ]);
    expect(routeShortcutAction('zoom to 80%')).toEqual([
      { type: 'SET_ZOOM', sheetName: undefined, zoomPercent: 80 },
    ]);
  });

  it('routes protect and unprotect commands', () => {
    expect(routeShortcutAction('protect this sheet')).toEqual([
      { type: 'PROTECT_SHEET', sheetName: undefined },
    ]);
    expect(routeShortcutAction('unprotect this sheet')).toEqual([
      { type: 'UNPROTECT_SHEET', sheetName: undefined },
    ]);
  });

  it('returns null for non-shortcut command', () => {
    expect(routeShortcutAction('calculate total sales for this month')).toBeNull();
  });

  it('returns null for conditional layout phrases', () => {
    expect(routeShortcutAction('hide rows where column A is blank')).toBeNull();
    expect(routeShortcutAction('freeze the rows with totals')).toBeNull();
    expect(routeShortcutAction('protect the sheet but allow column C')).toBeNull();
    expect(routeShortcutAction('hide rows where amount is 0')).toBeNull();
  });
});

describe('hasConditionalShortcutBlocker', () => {
  it('detects conditional keywords', () => {
    expect(hasConditionalShortcutBlocker('hide rows where amount is 0')).toBe(true);
    expect(hasConditionalShortcutBlocker('protect sheet but allow edits')).toBe(true);
    expect(hasConditionalShortcutBlocker('hide row 5')).toBe(false);
  });
});
