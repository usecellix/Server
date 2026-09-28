import { normalizeSingleAction } from '../src/agents/utils/normalize-executor-output.util';
import {
  ALL_SHEET_ACTION_TYPES,
  EXECUTOR_ADVERTISED_ACTION_TYPES,
} from '../src/excel-ai/types/action-catalog';

/**
 * TASKS.md #215 — "Hide the Narration column" produced "1 action, verified:
 * true" and then "Something went wrong applying this change". HIDE_COLUMN was
 * withheld from the Executor ("Tier 0 handles it" — only true when the user
 * names a column LETTER), so the model faked a hide with SET_COLUMN_WIDTH
 * width:0 carrying `columns: ["I"]`, which sanitizeAction dropped for want of
 * `col`. Two guards: advertise the real verbs, and convert the letter shapes
 * models actually emit instead of discarding them.
 */
describe('executor column/row action shapes (#215)', () => {
  describe('the verbs Tier 3 needs are advertised', () => {
    it.each([
      'HIDE_COLUMN',
      'UNHIDE_COLUMN',
      'HIDE_ROW',
      'UNHIDE_ROW',
      'SHOW_SHEET',
      'SET_SHEET_COLOR',
      'CLEAR_FORMAT',
      'UNMERGE_CELLS',
      'ADD_COMMENT',
    ])('%s is offered to the Executor', (type) => {
      expect(EXECUTOR_ADVERTISED_ACTION_TYPES).toContain(type);
    });

    it('keeps every type classified (the catalog stays exhaustive)', () => {
      expect(ALL_SHEET_ACTION_TYPES.length).toBeGreaterThan(EXECUTOR_ADVERTISED_ACTION_TYPES.length);
    });
  });

  describe('column letters are converted, not dropped', () => {
    it('converts a columns array to a 0-based index', () => {
      const action = normalizeSingleAction(
        { type: 'HIDE_COLUMN', sheetName: 'Purchase Register', columns: ['I'] },
        'Purchase Register',
      );
      expect(action).toMatchObject({ type: 'HIDE_COLUMN', col: 8, colCount: 1 });
    });

    it('spans a contiguous range from several letters', () => {
      const action = normalizeSingleAction(
        { type: 'HIDE_COLUMN', sheetName: 'Sheet1', columns: ['B', 'C', 'D'] },
        'Sheet1',
      );
      expect(action).toMatchObject({ col: 1, colCount: 3 });
    });

    it('accepts a singular column field', () => {
      // width kept above TASKS.md #265's 40pt readability floor — this test
      // is about column-letter conversion, not the clamp (covered separately
      // in normalize-executor-output.spec.ts).
      const action = normalizeSingleAction(
        { type: 'SET_COLUMN_WIDTH', sheetName: 'Sheet1', column: 'C', width: 90 },
        'Sheet1',
      );
      expect(action).toMatchObject({ col: 2, width: 90 });
    });

    it('leaves an explicit col index alone', () => {
      const action = normalizeSingleAction(
        { type: 'HIDE_COLUMN', sheetName: 'Sheet1', col: 3, colCount: 2, columns: ['Z'] },
        'Sheet1',
      );
      expect(action).toMatchObject({ col: 3, colCount: 2 });
    });

    it('ignores a header name that is not a column letter', () => {
      const action = normalizeSingleAction(
        { type: 'HIDE_COLUMN', sheetName: 'Sheet1', columns: ['Narration'] },
        'Sheet1',
      );
      expect(action?.col).toBeUndefined();
    });
  });
});
