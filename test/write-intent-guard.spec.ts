import {
  hasWriteIntent,
  isWorkbookScaffoldIntent,
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

  describe('workbook scaffold (multi-month + main dashboard)', () => {
    const repro =
      'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related things ,which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, total amount, source, payment status, bank account';

    it('detects write/scaffold intent for soft multi-sheet phrasing', () => {
      expect(isWorkbookScaffoldIntent(repro)).toBe(true);
      expect(hasWriteIntent(repro)).toBe(true);
    });

    it('does not treat bare which-question as read when full scaffold is present', () => {
      expect(
        hasWriteIntent(
          'need multiple sheets for all months and a main dashboard which all month sheets include total amount',
        ),
      ).toBe(true);
    });
  });

  // TASKS.md #235 — live-audit repros. Both were classified route=data by the
  // LLM router itself, reached applyWriteIntentGuard, and hasWriteIntent()
  // returned false for each — so the safety net that exists exactly for this
  // case let both through, and the request was answered with a chat question
  // ("want me to apply that change?") instead of a write-route preview.
  describe('conditional labeling verbs — "mark"/"flag" (#235)', () => {
    it.each([
      "If the GSTIN in column D is blank, mark a new Status column as 'Missing GSTIN'",
      "Add a column: if taxable amount is above 1 lakh mark 'High Value', else 'Standard'",
      'Flag rows where IGST is above zero',
    ])('detects write intent for %j', (message) => {
      expect(hasWriteIntent(message)).toBe(true);
    });
  });

  describe('"show/display only rows where" filter phrasing (#235)', () => {
    it.each([
      'Show only rows where the taxable amount is above 1 lakh',
      'Display only the rows where GSTIN is blank',
      'show only entries where the amount is above 50000',
    ])('detects write intent for %j', (message) => {
      expect(hasWriteIntent(message)).toBe(true);
    });

    it('leaves a plain read question alone', () => {
      expect(hasWriteIntent('How many rows where the taxable amount is above 1 lakh?')).toBe(
        false,
      );
    });
  });
});
