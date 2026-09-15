import {
  deriveTaxableValueFromSlabRow,
  extractRateFromColumnName,
  isRateSlabLayout,
} from './normalize-invoice';

describe('isRateSlabLayout', () => {
  it('detects a purchase rate-slab header layout (2+ slab columns)', () => {
    const headers = ['GSTIN', 'Invoice No', 'Invoice Date', 'Purchase@5%', 'Purchase@18%'];
    expect(isRateSlabLayout(headers)).toBe(true);
  });

  it('detects a sales rate-slab header layout including interstate columns', () => {
    const headers = ['GSTIN', 'Invoice No', 'Sales@18%', 'Sales Interstate@18%'];
    expect(isRateSlabLayout(headers)).toBe(true);
  });

  it('is false for a standard single Taxable Value column layout', () => {
    const headers = ['GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value'];
    expect(isRateSlabLayout(headers)).toBe(false);
  });

  it('is false when only one slab column is present', () => {
    const headers = ['GSTIN', 'Invoice No', 'Purchase@18%'];
    expect(isRateSlabLayout(headers)).toBe(false);
  });

  it('detects real-world header spacing variants ("Purchase @ 18%", "Purchase @ 28 %", etc.)', () => {
    const headers = [
      'GSTIN',
      'Invoice No',
      'Purchase @ 18%',
      'Purchase Interstate @18%',
      'Purchase @ 28 %',
      'Purchase@0%',
    ];
    expect(isRateSlabLayout(headers)).toBe(true);
  });
});

describe('extractRateFromColumnName', () => {
  it('extracts the rate from a standard slab column', () => {
    expect(extractRateFromColumnName('Purchase@18%')).toBe(18);
  });

  it('extracts the rate from an interstate slab column', () => {
    expect(extractRateFromColumnName('Purchase Interstate@18%')).toBe(18);
  });

  it('returns null when no rate can be parsed', () => {
    expect(extractRateFromColumnName('Taxable Value')).toBeNull();
  });

  it.each([
    ['Purchase @ 18%', 18],
    ['Purchase Interstate @18%', 18],
    ['Purchase @ 28 %', 28],
    ['Purchase@0%', 0],
  ])('extracts the rate from real-world spacing variant %s', (column, expected) => {
    expect(extractRateFromColumnName(column)).toBe(expected);
  });
});

describe('deriveTaxableValueFromSlabRow', () => {
  it('derives taxable value and rate from a populated standard-rate column', () => {
    const row = {
      GSTIN: '32AAAAA0000A1Z5',
      'Invoice No': 'INV-1',
      'Purchase@0%': '',
      'Purchase@5%': 0,
      'Purchase@18%': 12000,
      'Purchase Interstate@18%': '',
    };
    expect(deriveTaxableValueFromSlabRow(row)).toEqual({
      taxableValue: 12000,
      taxRatePercent: 18,
      sourceColumn: 'Purchase@18%',
    });
  });

  it('derives taxable value and rate from a populated interstate column', () => {
    const row = {
      GSTIN: '32AAAAA0000A1Z5',
      'Invoice No': 'INV-2',
      'Purchase@18%': '',
      'Purchase Interstate@18%': 5000,
    };
    expect(deriveTaxableValueFromSlabRow(row)).toEqual({
      taxableValue: 5000,
      taxRatePercent: 18,
      sourceColumn: 'Purchase Interstate@18%',
    });
  });

  it('returns null (not 0) when every slab column is blank or zero', () => {
    const row = {
      GSTIN: '32AAAAA0000A1Z5',
      'Invoice No': 'INV-3',
      'Purchase@0%': '',
      'Purchase@5%': 0,
      'Purchase@18%': null,
      'Purchase Interstate@18%': undefined,
    };
    expect(deriveTaxableValueFromSlabRow(row)).toBeNull();
  });

  it('matches slab column headers case-insensitively and with extra whitespace', () => {
    const row = { '  purchase@18%  ': 999 };
    expect(deriveTaxableValueFromSlabRow(row)).toEqual({
      taxableValue: 999,
      taxRatePercent: 18,
      sourceColumn: 'Purchase@18%',
    });
  });

  it('derives from real-world spaced headers ("Purchase @ 18%" etc.)', () => {
    const row = {
      GSTIN: '32AAACF3314R1ZO',
      'Purchase @ 0%': '',
      'Purchase @ 18%': 33495.78,
      'Purchase Interstate @18%': '',
      'Purchase @ 28 %': '',
    };
    expect(deriveTaxableValueFromSlabRow(row)).toEqual({
      taxableValue: 33495.78,
      taxRatePercent: 18,
      sourceColumn: 'Purchase@18%',
    });
  });
});

/**
 * Regression coverage for the "two non-blank rate-slab columns on one row" bug, audited
 * against every row of the customer's real "PR vs Missed PR.xlsx" Purchase register sheet
 * (338 rows). 9 rows had 2+ non-blank slab columns; values below are copied verbatim.
 * For 7 of them, exactly one candidate's implied tax (value × rate) matches the row's
 * actual CGST+SGST closely — that one must be picked, never the first one in priority
 * order (which was the bug: e.g. row 133 picked Purchase@12%=141.96 instead of the real
 * Purchase@18%=17166.11). The remaining 2 (rows 142 and 296) are genuine blended-rate
 * rows where no single candidate's implied tax is close — those must be flagged
 * ambiguous, never guessed.
 */
describe('deriveTaxableValueFromSlabRow — multi-slab disambiguation (real audited rows)', () => {
  it.each([
    // [excel row, candidates, actualTax (CGST+SGST), expected picked value, expected rate]
    [133, { 'Purchase@12%': 141.96, 'Purchase @ 18%': 17166.11 }, 3106.94, 17166.11, 18],
    [140, { 'Purchase@0%': 3816, 'Purchase @ 18%': 28380.51 }, 5108.5, 28380.51, 18],
    [302, { 'Purchase@12%': 152.7, 'Purchase @ 18%': 32287.37 }, 5830.04, 32287.37, 18],
    [303, { 'Purchase@0%': 2550, 'Purchase@5%': 485.7, 'Purchase@12%': 0.01, 'Purchase @ 18%': 35542.91 }, 6422, 35542.91, 18],
    [313, { 'Purchase@0%': 340, 'Purchase @ 18%': 35076.53 }, 6313.78, 35076.53, 18],
    [316, { 'Purchase@0%': 610, 'Purchase @ 18%': 19517.29 }, 3513.14, 19517.29, 18],
    [322, { 'Purchase@0%': 177, 'Purchase @ 18%': 8677.17 }, 1561.9, 8677.17, 18],
  ])('row %i: picks the candidate whose implied tax matches actual tax ₹%s, never the first in priority order', (
    _excelRow,
    candidates,
    actualTax,
    expectedValue,
    expectedRate,
  ) => {
    const result = deriveTaxableValueFromSlabRow(candidates as Record<string, unknown>, actualTax as number);
    expect(result).not.toBeNull();
    expect(result && 'ambiguous' in result).toBe(false);
    expect((result as { taxableValue: number }).taxableValue).toBe(expectedValue);
    expect((result as { taxRatePercent: number | null }).taxRatePercent).toBe(expectedRate);
  });

  it('row 133: without the actual-tax fix, the bug would have picked Purchase@12%=141.96 (wrong)', () => {
    // Confirms the fix actually changes behavior versus first-in-priority-order selection —
    // Purchase@12% comes before Purchase@18% in RATE_SLAB_COLUMNS.
    const row = { 'Purchase@12%': 141.96, 'Purchase @ 18%': 17166.11 };
    const result = deriveTaxableValueFromSlabRow(row, 3106.94);
    expect((result as { taxableValue: number }).taxableValue).not.toBe(141.96);
    expect((result as { taxableValue: number }).taxableValue).toBe(17166.11);
  });

  it.each([
    [142, { 'Purchase@12%': 4017.86, 'Purchase @ 18%': 4555.08 }, 1302.06],
    [296, { 'Purchase@0%': 4200, 'Purchase@5%': 2971.2, 'Purchase @ 18%': 13542.98 }, 2586.3],
  ])(
    'row %i: flags ambiguous_rate_slab instead of guessing when no candidate is a close tax match',
    (_excelRow, candidates, actualTax) => {
      const result = deriveTaxableValueFromSlabRow(candidates as Record<string, unknown>, actualTax as number);
      expect(result).not.toBeNull();
      expect(result && 'ambiguous' in result).toBe(true);
      const ambiguous = result as { ambiguous: true; candidates: unknown[]; detail: string };
      expect(ambiguous.candidates.length).toBeGreaterThan(1);
      expect(ambiguous.detail).toMatch(/multiple rate-slab columns/i);
    },
  );

  it('still resolves normally (no actualTax needed) when exactly one slab column is populated', () => {
    const row = { 'Purchase@18%': 5000, 'Purchase@12%': '' };
    expect(deriveTaxableValueFromSlabRow(row)).toEqual({
      taxableValue: 5000,
      taxRatePercent: 18,
      sourceColumn: 'Purchase@18%',
    });
  });

  it('flags ambiguous (rather than defaulting to the first candidate) when actualTax is not provided at all', () => {
    const row = { 'Purchase@12%': 141.96, 'Purchase @ 18%': 17166.11 };
    const result = deriveTaxableValueFromSlabRow(row);
    expect(result && 'ambiguous' in result).toBe(true);
  });
});
