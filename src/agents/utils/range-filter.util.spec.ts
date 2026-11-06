import { applyFilterOperator, filterDataRows, findMatchingRowOffsets } from './range-filter.util';

const f = (operator: string, value: string | number = '', extra: Record<string, unknown> = {}) =>
  ({ column: 'A', operator, value, ...extra }) as never;

describe('P1.7 server range filter (shadow-workbook mirror of the add-in)', () => {
  it('blank / not blank, whitespace only with the flag', () => {
    expect(applyFilterOperator('', f('isBlank'))).toBe(true);
    expect(applyFilterOperator(' ', f('isBlank'))).toBe(false);
    expect(applyFilterOperator(' ', f('isBlank', 'trimSpaces'))).toBe(true);
    expect(applyFilterOperator('x', f('isNotBlank'))).toBe(true);
  });

  it('text numbers', () => {
    expect(applyFilterOperator('1,234', f('isTextNumber'))).toBe(true);
    expect(applyFilterOperator(1234, f('isTextNumber'))).toBe(false);
  });

  it('invalid GSTIN: typo flagged, valid and blank not', () => {
    expect(applyFilterOperator('27AABCA1234F1Z7', f('isInvalidGstin'))).toBe(true);
    expect(applyFilterOperator('27AABCA1234F1Z9', f('isInvalidGstin'))).toBe(false);
    expect(applyFilterOperator('', f('isInvalidGstin'))).toBe(false);
  });

  it('olderThan with asOn', () => {
    expect(applyFilterOperator('01-01-2026', f('olderThan', 90, { asOn: '2026-04-30' }))).toBe(true);
    expect(applyFilterOperator('15-04-2026', f('olderThan', 90, { asOn: '2026-04-30' }))).toBe(false);
  });

  it('isDuplicate keep first / last through findMatchingRowOffsets and filterDataRows', () => {
    const rows = [['Inv'], ['A'], ['B'], ['a'], ['']];
    expect(findMatchingRowOffsets(rows, true, { column: 'Inv', operator: 'isDuplicate', value: '' } as never)).toEqual([3]);
    expect(findMatchingRowOffsets(rows, true, { column: 'Inv', operator: 'isDuplicate', value: '', keep: 'last' } as never)).toEqual([1]);
    expect(filterDataRows(rows, true, { column: 'Inv', operator: 'isDuplicate', value: '' } as never).filteredRows).toEqual([['a']]);
  });
});
