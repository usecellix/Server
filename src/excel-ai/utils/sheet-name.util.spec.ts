import {
  extractSheetNameFromPrompt,
  MAX_EXCEL_SHEET_NAME_LENGTH,
  sanitizeExcelSheetName,
  stripIllegalSheetNameChars,
} from './sheet-name.util';

describe('sanitizeExcelSheetName', () => {
  it.each([
    ['\\', 'A\\B', 'A B'],
    ['/', 'A/B', 'A B'],
    ['?', 'A?B', 'A B'],
    ['*', 'A*B', 'A B'],
    ['[', 'A[B', 'A B'],
    [']', 'A]B', 'A B'],
    [':', 'A:B', 'A B'],
  ])('strips the illegal character %s', (_label, input, expected) => {
    expect(sanitizeExcelSheetName(input)).toBe(expected);
  });

  it('strips every illegal character in one name and collapses the resulting whitespace', () => {
    expect(sanitizeExcelSheetName('A\\B/C?D*E[F]G:H')).toBe('A B C D E F G H');
  });

  it('the exact regression case: "Missed vs GSTR-2B / GSTR-2A" (dual-portal-source sheet name)', () => {
    expect(sanitizeExcelSheetName('Missed vs GSTR-2B / GSTR-2A')).toBe(
      'Missed vs GSTR-2B GSTR-2A',
    );
    expect(sanitizeExcelSheetName('Missed vs GSTR-2B / GSTR-2A')).not.toMatch(/[\\/?*[\]:]/);
  });

  it('enforces the 31-character Excel sheet-name limit', () => {
    const long = 'A'.repeat(40);
    const result = sanitizeExcelSheetName(long);
    expect(result.length).toBe(MAX_EXCEL_SHEET_NAME_LENGTH);
    expect(result).toBe('A'.repeat(31));
  });

  it('strips illegal characters AND truncates together, in one pass', () => {
    const long = `${'A/'.repeat(20)}B`; // 41 chars raw, alternating illegal char
    const result = sanitizeExcelSheetName(long);
    expect(result.length).toBeLessThanOrEqual(MAX_EXCEL_SHEET_NAME_LENGTH);
    expect(result).not.toMatch(/[\\/?*[\]:]/);
  });

  it('falls back to the given fallback (default "Sheet") when the sanitized result is empty', () => {
    expect(sanitizeExcelSheetName('///')).toBe('Sheet');
    expect(sanitizeExcelSheetName('***', 'Untitled')).toBe('Untitled');
    expect(sanitizeExcelSheetName('   ')).toBe('Sheet');
  });

  it('leaves an already-valid name unchanged', () => {
    expect(sanitizeExcelSheetName('Missed vs GSTR-2B')).toBe('Missed vs GSTR-2B');
  });
});

describe('stripIllegalSheetNameChars', () => {
  it('strips illegal characters without truncating to 31 chars', () => {
    const long = 'A'.repeat(40);
    expect(stripIllegalSheetNameChars(long)).toBe(long);
    expect(stripIllegalSheetNameChars(long).length).toBe(40);
  });

  it('strips illegal characters without falling back on an empty/all-illegal result', () => {
    expect(stripIllegalSheetNameChars('///')).toBe('');
    expect(stripIllegalSheetNameChars('')).toBe('');
  });

  it('lets a caller reserve room for a suffix before truncating the combined name', () => {
    const suffix = ' — Books Only';
    const base = stripIllegalSheetNameChars('Missed vs GSTR-2B / GSTR-2A').slice(
      0,
      MAX_EXCEL_SHEET_NAME_LENGTH - suffix.length,
    );
    const combined = `${base}${suffix}`;
    expect(combined.length).toBeLessThanOrEqual(MAX_EXCEL_SHEET_NAME_LENGTH);
    expect(combined.endsWith(suffix)).toBe(true);
    expect(combined).not.toMatch(/[\\/?*[\]:]/);
  });
});

describe('extractSheetNameFromPrompt — unaffected by the sanitizer refactor', () => {
  it('still extracts and sanitizes a quoted sheet name', () => {
    expect(extractSheetNameFromPrompt('create a sheet named "Bad/Name?"')).toBe('Bad Name');
  });
});
