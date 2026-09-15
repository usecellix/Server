import { routeShortcutAction } from '../src/excel-ai/utils/shortcut-router.util';

/**
 * TASKS.md #216 — the shortcut lane resolved targets by regex alone, so phrases
 * straight out of `cellix-basic-usecases.html` hit the wrong column, the wrong
 * sheet, or the wrong feature entirely. Falling through to the LLM (null) is a
 * correct outcome here: it can resolve a header name against real headers.
 */
describe('shortcut router target resolution (#216)', () => {
  const active = 'Purchase Register';

  describe('column letters vs header names', () => {
    it.each([
      'Hide column Amount',
      'Show column totals',
      'Hide the column called Notes',
      'Set column width of column B to 20',
    ])('does not invent a column index for %j', (message) => {
      const actions = routeShortcutAction(message, active) ?? [];
      for (const action of actions) {
        if (typeof action.col === 'number') expect(action.col).toBeLessThanOrEqual(16383);
      }
      // A header name has no business resolving to a letter-derived index.
      expect(actions.some((a) => typeof a.col === 'number' && a.col > 25)).toBe(false);
    });

    it('still resolves real column letters', () => {
      expect(routeShortcutAction('Hide column D', active)).toEqual([
        { type: 'HIDE_COLUMN', sheetName: active, col: 3, colCount: 1 },
      ]);
      expect(routeShortcutAction('Hide columns B to D', active)).toEqual([
        { type: 'HIDE_COLUMN', sheetName: active, col: 1, colCount: 3 },
      ]);
      expect(routeShortcutAction('Unhide column C', active)).toEqual([
        { type: 'UNHIDE_COLUMN', sheetName: active, col: 2, colCount: 1 },
      ]);
    });
  });

  describe('sheet targets', () => {
    it('strips the called/named lead-in', () => {
      expect(routeShortcutAction('Hide the sheet called Working', active)).toEqual([
        { type: 'HIDE_SHEET', sheetName: 'Working' },
      ]);
    });

    it('understands "X sheet" word order', () => {
      expect(routeShortcutAction('Hide the Working sheet', active)).toEqual([
        { type: 'HIDE_SHEET', sheetName: 'Working' },
      ]);
    });

    it('hides the active sheet only when the user said so', () => {
      expect(routeShortcutAction('Hide this sheet', active)).toEqual([
        { type: 'HIDE_SHEET', sheetName: active },
      ]);
    });

    it('does not treat a listing question as an unhide', () => {
      expect(routeShortcutAction('Show the sheets in this workbook', active)).toBeNull();
      expect(routeShortcutAction('Show sheet list', active)).toBeNull();
    });

    it('colours the sheet the user named, not the active one', () => {
      expect(routeShortcutAction('Colour the Summary tab blue', active)).toEqual([
        { type: 'SET_SHEET_COLOR', sheetName: 'Summary', color: '#0000FF' },
      ]);
    });
  });

  describe('freeze vs protect', () => {
    it('leaves "lock the header row so nobody edits it" to the LLM', () => {
      expect(routeShortcutAction('Lock the header row so nobody edits it', active)).toBeNull();
    });

    it('still freezes a plain freeze request', () => {
      expect(routeShortcutAction('Freeze the top row', active)).toEqual([
        { type: 'FREEZE_PANES', sheetName: active, freezeRows: 1, freezeColumns: 0 },
      ]);
    });
  });

  describe('row and column sizing carries the count', () => {
    it('sizes every row in the range', () => {
      expect(routeShortcutAction('Set row height of rows 2 to 5 to 25', active)).toEqual([
        { type: 'SET_ROW_HEIGHT', sheetName: active, row: 1, rowCount: 4, height: 25 },
      ]);
    });
  });
});
