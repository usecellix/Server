import { applyFilterOperator, filterDataRows, buildOutputRows } from '../src/agents/utils/range-filter.util';

describe('range-filter.util', () => {
  it('applies equals / contains / notEquals case-insensitively', () => {
    expect(
      applyFilterOperator('Pending', { column: 'Status', operator: 'equals', value: 'pending' }),
    ).toBe(true);
    expect(
      applyFilterOperator('Payment Pending', {
        column: 'Status',
        operator: 'contains',
        value: 'pending',
      }),
    ).toBe(true);
    expect(
      applyFilterOperator('Paid', { column: 'Status', operator: 'notEquals', value: 'Pending' }),
    ).toBe(true);
  });

  it('applies numeric greaterThan / lessThan', () => {
    expect(
      applyFilterOperator(100, { column: 'Amt', operator: 'greaterThan', value: 50 }),
    ).toBe(true);
    expect(
      applyFilterOperator(10, { column: 'Amt', operator: 'lessThan', value: 50 }),
    ).toBe(true);
    expect(
      applyFilterOperator('x', { column: 'Amt', operator: 'greaterThan', value: 50 }),
    ).toBe(false);
  });

  it('applies lengthEquals / lengthNotEquals and regex operators', () => {
    expect(
      applyFilterOperator('ABCDE1234567890', {
        column: 'GSTIN',
        operator: 'lengthEquals',
        value: 15,
      }),
    ).toBe(true);
    expect(
      applyFilterOperator('SHORT', {
        column: 'GSTIN',
        operator: 'lengthNotEquals',
        value: 15,
      }),
    ).toBe(true);
    expect(
      applyFilterOperator('ABCDE1234567890', {
        column: 'GSTIN',
        operator: 'matchesRegex',
        value: '^[A-Za-z0-9]{15}$',
      }),
    ).toBe(true);
    expect(
      applyFilterOperator('BAD-GSTIN!!', {
        column: 'GSTIN',
        operator: 'notMatchesRegex',
        value: '^[A-Za-z0-9]{15}$',
      }),
    ).toBe(true);
    expect(
      applyFilterOperator('ABCDE1234567890', {
        column: 'GSTIN',
        operator: 'notMatchesRegex',
        value: '^[A-Za-z0-9]{15}$',
      }),
    ).toBe(false);
  });

  it('filters data rows by header column', () => {
    const rows = [
      ['Name', 'Payment Status'],
      ['A', 'Pending'],
      ['B', 'Paid'],
      ['C', 'Pending'],
    ];
    const { headerRow, filteredRows } = filterDataRows(rows, true, {
      column: 'Payment Status',
      operator: 'equals',
      value: 'Pending',
    });
    expect(headerRow).toEqual(['Name', 'Payment Status']);
    expect(filteredRows).toEqual([
      ['A', 'Pending'],
      ['C', 'Pending'],
    ]);
    expect(buildOutputRows(headerRow, filteredRows)).toHaveLength(3);
  });

  it('throws when filter column is missing', () => {
    expect(() =>
      filterDataRows([['A', 'B'], [1, 2]], true, {
        column: 'Missing',
        operator: 'equals',
        value: 'x',
      }),
    ).toThrow(/not found/);
  });
});
