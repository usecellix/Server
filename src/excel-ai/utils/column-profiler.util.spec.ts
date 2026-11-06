import {
  buildProfileReport,
  dateKindOf,
  detectProfileIntent,
  isTextNumber,
  profileColumns,
  profileOnlyColumnsFromMessage,
} from './column-profiler.util';

// Same data as T_Profile in CELLIX_TEST_DATA.xlsx
const T_PROFILE: unknown[][] = [
  ['TextNumber', 'DrCrFormat', 'TextDate', 'Spaces', 'MixedCase'],
  ['1,234', 1500, '01-04-2025', 'ABC ', 'abc'],
  ['5,000.50', 2500, '1.4.25', ' XYZ', 'ABC'],
  ['\u20B9 7,500', 3500, '20250401', 'Mehta\u00A0Sons', 'Abc'],
  ['(2,500)', 4500, '13-04-2025', 'Clean', 'aBc'],
];
const FORMATS = [
  undefined,
  [
    { format: '#,##0.00\ "Cr"', count: 2 },
    { format: '#,##0.00\ "Dr"', count: 2 },
  ],
];

describe('P1.5 column profiler', () => {
  it('text-stored numbers: commas, rupee sign, brackets; not dates/invoice text', () => {
    for (const v of ['1,234', '5,000.50', '\u20B9 7,500', '(2,500)', '-45', '12 Cr'])
      expect(isTextNumber(v)).toBe(true);
    for (const v of ['01-04-2025', '20250401', 'INV-233', 'abc', '', 1234, null])
      expect(isTextNumber(v)).toBe(false);
  });

  it('classifies text date layouts', () => {
    expect(dateKindOf('01-04-2025')).toBe('dd-mm-yyyy');
    expect(dateKindOf('1.4.25')).toBe('d.m.yy');
    expect(dateKindOf('20250401')).toBe('yyyymmdd');
    expect(dateKindOf('2025-04-01')).toBe('yyyy-mm-dd');
    expect(dateKindOf('01-Apr-2025')).toBe('dd-Mon-yyyy');
    expect(dateKindOf('Apr 1, 2025')).toBe('Mon dd, yyyy');
    expect(dateKindOf('hello')).toBeNull();
  });

  describe('T_Profile fixture', () => {
    const p = profileColumns(T_PROFILE, { headerRowIndex: 0, formatCounts: FORMATS });
    it('A: 4 text-stored numbers', () => expect(p[0].textNumbers).toBe(4));
    it('B: Dr/Cr found in the number format (2 Cr, 2 Dr)', () => {
      expect(p[1].crFormatCells).toBe(2);
      expect(p[1].drFormatCells).toBe(2);
    });
    it('C: 3 date formats', () => {
      expect(p[2].dateKinds.sort()).toEqual(['d.m.yy', 'dd-mm-yyyy', 'yyyymmdd'].sort());
      expect(p[2].issues.join(' ')).toMatch(/3 different text date formats/);
    });
    it('D: 2 space-padded + 1 NBSP', () => {
      expect(p[3].spacePadded).toBe(2);
      expect(p[3].nbsp).toBe(1);
    });
    it('E: mixed case', () => expect(p[4].mixedCase).toBe(true));
  });

  it('does not mark a name column mixed-case for a few title vs upper values', () => {
    const rows: unknown[][] = [['Party'], ['ABC Traders'], ['Mehta Sons'], ['Shah Bros'], ['Kumar Agency'], ['Patel Stores'], ['Rao Co'], ['Jain Mart']];
    expect(profileColumns(rows)[0].mixedCase).toBe(false);
  });

  it('does not flag an ID column with one stray text number', () => {
    const rows: unknown[][] = [['Ref'], ['INV-1'], ['INV-2'], ['0233'], ['INV-4']];
    expect(profileColumns(rows)[0].issues.join()).not.toMatch(/text-stored/);
  });

  it('detects hidden control characters and tab', () => {
    const rows: unknown[][] = [['A'], ['Kumar\tAgency'], ['ok'], ['x\u200By']];
    expect(profileColumns(rows)[0].nonPrinting).toBe(2);
  });

  it('counts blanks inside the data, not trailing empty rows', () => {
    const rows: unknown[][] = [['A'], ['x'], [''], ['y'], [''], ['']];
    expect(profileColumns(rows)[0].blanks).toBe(1);
  });

  it('only the requested column', () => {
    const p = profileColumns(T_PROFILE, { onlyColumns: [3] });
    expect(p).toHaveLength(1);
    expect(p[0].letter).toBe('D');
  });

  it('is read-only (does not mutate input)', () => {
    const copy = JSON.stringify(T_PROFILE);
    profileColumns(T_PROFILE, { formatCounts: FORMATS });
    expect(JSON.stringify(T_PROFILE)).toBe(copy);
  });

  describe('intent', () => {
    it.each([
      'Check the data quality of this sheet',
      "What's wrong with these columns?",
      "what's wrong with column D?",
      'health check this sheet',
    ])('detects: %s', (m) => expect(detectProfileIntent(m)).toBe(true));
    it.each([
      "What's wrong with this formula in column D?",
      'Why is VLOOKUP giving #N/A',
      'Total of column D',
      'Validate the GSTINs in column B',
    ])('ignores: %s', (m) => expect(detectProfileIntent(m)).toBe(false));
    it('column scope', () => {
      expect(profileOnlyColumnsFromMessage("what's wrong with column D?")).toEqual([3]);
      expect(profileOnlyColumnsFromMessage('check the data quality of this sheet')).toBeUndefined();
    });
  });

  it('report lists every flagged column and says nothing was changed', () => {
    const p = profileColumns(T_PROFILE, { formatCounts: FORMATS });
    const { answer, flaggedColumns } = buildProfileReport('T_Profile', p);
    expect(flaggedColumns).toBe(5);
    expect(answer).toMatch(/Nothing was changed/);
    expect(answer).toMatch(/Column D/);
  });

  it('report says clean when nothing is wrong', () => {
    const p = profileColumns([['A'], ['x'], ['y']]);
    expect(buildProfileReport('S', p).answer).toMatch(/looks clean/);
  });
});
