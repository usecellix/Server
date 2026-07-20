import {
  hasWriteIntent,
  WRITE_INTENT_CATALOG_VERBS,
} from '../src/excel-ai/utils/write-intent-guard.util';

describe('hasWriteIntent', () => {
  describe('catalog verbs (Spec 01 / Spec 11)', () => {
    it.each(WRITE_INTENT_CATALOG_VERBS)('returns true for verb "%s"', (verb) => {
      expect(hasWriteIntent(`please ${verb} the sheet`)).toBe(true);
    });
  });

  describe('exact repro — sort with Total Amount (data-lane trap)', () => {
    const repro = 'sort the sheet based on Total Amount descending';

    it('detects write intent', () => {
      expect(hasWriteIntent(repro)).toBe(true);
    });

    it('is stable across repeated calls (no non-determinism)', () => {
      for (let i = 0; i < 10; i += 1) {
        expect(hasWriteIntent(repro)).toBe(true);
      }
    });
  });

  describe('Spec 01 catalog phrasings that imply mutation', () => {
    const writePhrases = [
      'bold cells A1 to C1',
      'freeze top row',
      'hide column F',
      'insert a row above row 10',
      'delete column C',
      'sort column B descending by value',
      'filter the data by Status',
      'find and replace ABC with XYZ',
      'highlight cells greater than 100',
      'fill down the formula in column D',
      'create a pivot table from this data',
      'build pivot table by Region',
      'create a chart from sales data',
      'add a bar graph for Q1 revenue',
      'add data validation dropdown in column B',
      'sort by Amount ascending',
      'apply formatting across all sheets',
    ];

    it.each(writePhrases)('returns true for "%s"', (message) => {
      expect(hasWriteIntent(message)).toBe(true);
    });
  });

  describe('read-intent overrides', () => {
    it('keeps "what would happen if I sorted" as read-only', () => {
      expect(hasWriteIntent('what would happen if I sorted this sheet')).toBe(false);
    });

    it('keeps how-many / show-me / explain questions read-only', () => {
      expect(hasWriteIntent('how many rows would delete remove')).toBe(false);
      expect(hasWriteIntent('show me the total amount')).toBe(false);
      expect(hasWriteIntent('explain how to sort a column')).toBe(false);
      expect(hasWriteIntent('why is the color wrong')).toBe(false);
      expect(hasWriteIntent('can you tell me which column to sort')).toBe(false);
      expect(hasWriteIntent('is there a way to filter this')).toBe(false);
    });

    it('still treats compound ask-then-mutate as write', () => {
      expect(hasWriteIntent('what is the total and then sort by Amount')).toBe(true);
      expect(hasWriteIntent('show me blanks then delete them')).toBe(true);
    });
  });

  describe('pure read queries without write verbs', () => {
    it('returns false for aggregation/find without mutation verbs', () => {
      expect(hasWriteIntent('what is the sum of column B')).toBe(false);
      expect(hasWriteIntent('find invoice INV-100')).toBe(false);
      expect(hasWriteIntent('how many invoices are there')).toBe(false);
    });
  });
});
