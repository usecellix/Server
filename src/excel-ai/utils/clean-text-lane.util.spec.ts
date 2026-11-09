import {
  buildCleanTextPlan,
  cleanColumnsFromMessage,
  cleanText,
  detectCleanTextIntent,
  showSpaces,
} from './clean-text-lane.util';

// Same data as T_Clean in CELLIX_TEST_DATA.xlsx (row 8 column A is a formula cell).
const T_CLEAN: unknown[][] = [
  ['Party', 'Amount', 'Date', 'Note'],
  ['  ABC Traders', 1234.5, 46113, 'leading spaces'],
  ['XYZ Pvt  Ltd', 2000, 46114, 'double space'],
  ['Mehta\u00A0& Sons', 3000, 46115, 'NBSP (CHAR 160)'],
  ['Kumar\tAgency', 4000, 46116, 'tab character'],
  ['Shah Bros  ', 5000, 46117, 'trailing spaces'],
  ['Clean Name', 6000, 46118, 'already clean'],
  ['  Formula Cell  ', 7000, 46119, 'formula - must be skipped'],
];
const FORMULA_ROWS = [[7]]; // column A, row index 7

const opsOf = (plan: ReturnType<typeof buildCleanTextPlan>) =>
  Object.fromEntries(((plan.actions![0] as { operations: Array<{ address: string; value: unknown }> }).operations).map((o) => [o.address, o.value]));

describe('P1.9 clean text lane', () => {
  describe('intent', () => {
    it.each([
      'Clean the spaces and hidden characters in column A',
      'Trim spaces in column A',
      'trim column B',
      'Remove extra spaces from column C',
      'remove non-printing characters in column A',
      'Remove non-breaking spaces',
    ])('detects: %s', (m) => expect(detectCleanTextIntent(m)).toBe(true));

    it.each([
      'Delete blank rows',
      'Remove duplicate rows',
      'Validate the GSTINs in column B',
      'Total of column D',
      'Remove the Remarks column',
      '',
    ])('ignores: %s', (m) => expect(detectCleanTextIntent(m)).toBe(false));
  });

  it('parses column scope', () => {
    expect(cleanColumnsFromMessage('clean spaces in column A')).toEqual([0]);
    expect(cleanColumnsFromMessage('trim columns A and C')).toEqual([0, 2]);
    expect(cleanColumnsFromMessage('trim spaces')).toBeUndefined();
  });

  describe('cleanText', () => {
    it.each([
      ['  ABC Traders', 'ABC Traders'],
      ['XYZ Pvt  Ltd', 'XYZ Pvt Ltd'],
      ['Mehta\u00A0& Sons', 'Mehta & Sons'],
      ['Kumar\tAgency', 'Kumar Agency'],
      ['Shah Bros  ', 'Shah Bros'],
      ['a\u200Bb', 'ab'],
      ['x\u0007y', 'xy'],
      ['Clean Name', 'Clean Name'],
    ])('%j → %j', (input, out) => expect(cleanText(input)).toBe(out));

    it('only applies the requested ops', () => {
      expect(cleanText('  a  b ', ['trim'])).toBe('a  b');
    });
  });

  it('shows invisible characters', () => {
    expect(showSpaces('  A b')).toBe('··A·b');
    expect(showSpaces('a\u00A0b')).toBe('a⍽b');
    expect(showSpaces('a\tb')).toBe('a→b');
  });

  describe('T_Clean fixture, column A', () => {
    const plan = buildCleanTextPlan(T_CLEAN, 'T_Clean', 'Clean the spaces and hidden characters in column A', {
      headerRowIndex: 0,
      formulaRows: FORMULA_ROWS,
    });

    it('5 to clean, 1 already clean, 1 formula skipped', () => {
      expect(plan.kind).toBe('actions');
      expect(plan.counts).toMatchObject({ toClean: 5, alreadyClean: 1, formulaSkipped: 1, keptAsText: 0 });
    });

    it('writes exactly the cleaned text, nothing else', () => {
      expect(opsOf(plan)).toEqual({
        A2: 'ABC Traders',
        A3: 'XYZ Pvt Ltd',
        A4: 'Mehta & Sons',
        A5: 'Kumar Agency',
        A6: 'Shah Bros',
      });
    });

    it('never touches numbers, dates, other columns or the formula cell', () => {
      const addrs = Object.keys(opsOf(plan));
      expect(addrs.some((a) => /^[BCD]/.test(a))).toBe(false);
      expect(addrs).not.toContain('A8');
    });

    it('is one BATCH_SET on the sheet (preview → Accept → undo)', () => {
      expect(plan.actions).toHaveLength(1);
      expect(plan.actions![0]).toMatchObject({ type: 'BATCH_SET', sheetName: 'T_Clean' });
    });

    it('preview text reports counts and shows · for spaces', () => {
      expect(plan.answer).toMatch(/5 cells to clean/);
      expect(plan.answer).toMatch(/1 already clean/);
      expect(plan.answer).toMatch(/1 formula cell skipped/);
      expect(plan.answer).toContain('A2: ··ABC·Traders → ABC·Traders');
    });
  });

  it('without a column, scans every text column but leaves numbers and dates alone', () => {
    const plan = buildCleanTextPlan(T_CLEAN, 'T_Clean', 'Clean the spaces in this sheet', { formulaRows: FORMULA_ROWS });
    expect(Object.keys(opsOf(plan)).sort()).toEqual(['A2', 'A3', 'A4', 'A5', 'A6']);
  });

  it('does not turn cleaned text into numbers (leading zeros survive)', () => {
    const sheet: unknown[][] = [['Code'], [' 00123 '], ['A 1 ']];
    const plan = buildCleanTextPlan(sheet, 'S', 'trim spaces in column A', {});
    expect(plan.counts).toMatchObject({ toClean: 1, keptAsText: 1 });
    expect(opsOf(plan)).toEqual({ A3: 'A 1' });
    expect(plan.answer).toMatch(/left unchanged because cleaning would make Excel turn/);
  });

  it('answers (no write) when everything is already clean', () => {
    const plan = buildCleanTextPlan([['A'], ['x'], ['y z']], 'S', 'trim spaces in column A', {});
    expect(plan.kind).toBe('answer');
    expect(plan.answer).toMatch(/Nothing to clean/);
  });
});
