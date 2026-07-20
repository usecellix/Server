import { buildAggregateTable } from '../src/agents/utils/aggregate-table.util';

describe('buildAggregateTable', () => {
  const rows = [
    ['Supplier', 'Total Amount', 'Qty'],
    ['Acme', 100, 2],
    ['Beta', 50, 1],
    ['Acme', 25, 3],
    ['Gamma', 200, 4],
  ];

  it('groups and sums with topN desc sort', () => {
    const table = buildAggregateTable({
      rows,
      hasHeaders: true,
      groupByColumn: 'Supplier',
      aggregations: [{ column: 'Total Amount', fn: 'sum', outputLabel: 'Total Spend' }],
      sortBy: { column: 'Total Spend', direction: 'desc' },
      topN: 2,
    });

    expect(table[0]).toEqual(['Supplier', 'Total Spend']);
    expect(table[1]).toEqual(['Gamma', 200]);
    expect(table[2]).toEqual(['Acme', 125]);
    expect(table).toHaveLength(3);
  });

  it('supports count and average', () => {
    const table = buildAggregateTable({
      rows,
      hasHeaders: true,
      groupByColumn: 'Supplier',
      aggregations: [
        { column: 'Qty', fn: 'count', outputLabel: 'Orders' },
        { column: 'Total Amount', fn: 'average', outputLabel: 'Avg' },
      ],
    });

    const acme = table.find((r) => r[0] === 'Acme');
    expect(acme?.[1]).toBe(2);
    expect(acme?.[2]).toBe(62.5);
  });

  it('throws when group-by column is missing', () => {
    expect(() =>
      buildAggregateTable({
        rows,
        hasHeaders: true,
        groupByColumn: 'Missing',
        aggregations: [{ column: 'Total Amount', fn: 'sum', outputLabel: 'Spend' }],
      }),
    ).toThrow(/not found/);
  });
});
