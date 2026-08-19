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

  it('groups by month transform from a date column', () => {
    const dated = [
      ['Date', 'Amount'],
      ['2024-01-15', 10],
      ['2024-01-20', 5],
      ['2024-02-01', 20],
      ['2024-03-10', 7],
    ];
    const table = buildAggregateTable({
      rows: dated,
      hasHeaders: true,
      groupByColumn: 'Date',
      groupByTransform: 'month',
      aggregations: [{ column: 'Amount', fn: 'sum', outputLabel: 'Total' }],
    });

    expect(table[0]).toEqual(['Month', 'Total']);
    const jan = table.find((r) => r[0] === 'Jan');
    const feb = table.find((r) => r[0] === 'Feb');
    const mar = table.find((r) => r[0] === 'Mar');
    expect(jan?.[1]).toBe(15);
    expect(feb?.[1]).toBe(20);
    expect(mar?.[1]).toBe(7);
  });

  describe('fn: "first" — passthrough label column 1:1 with the group key', () => {
    // Regression: AGGREGATE_TABLE only ever supported one groupByColumn, with no
    // way to carry a second identity column through untouched. A request like
    // "GSTIN-wise summary ... for each supplier" needs GSTIN as the group key
    // AND Supplier Name in the output — previously inexpressible, so the
    // Executor either omitted groupByColumn or invented a second one, and the
    // action silently no-op'd (virtualAggregateTable returns early when
    // groupByColumn is falsy), surfacing as "Aggregation is missing required
    // group-by fields" after retries were exhausted.
    const gstRows = [
      ['GSTIN', 'Supplier Name', 'Taxable Value', 'Tax Amount', 'Invoice Value'],
      ['29AABCT1332L1Z5', 'Acme Traders', 10000, 1800, 11800],
      ['29AABCT1332L1Z5', 'Acme Traders', 5000, 900, 5900],
      ['07AAACB2230P1Z1', 'Beta Supplies', 20000, 3600, 23600],
    ];

    it('groups by GSTIN and carries Supplier Name through as fn: "first"', () => {
      const table = buildAggregateTable({
        rows: gstRows,
        hasHeaders: true,
        groupByColumn: 'GSTIN',
        aggregations: [
          { column: 'Supplier Name', fn: 'first', outputLabel: 'Supplier Name' },
          { column: 'Taxable Value', fn: 'sum', outputLabel: 'Total Taxable Value' },
          { column: 'Tax Amount', fn: 'sum', outputLabel: 'Total Tax Amount' },
          { column: 'Invoice Value', fn: 'sum', outputLabel: 'Total Invoice Value' },
        ],
      });

      expect(table[0]).toEqual([
        'GSTIN',
        'Supplier Name',
        'Total Taxable Value',
        'Total Tax Amount',
        'Total Invoice Value',
      ]);
      const acme = table.find((r) => r[0] === '29AABCT1332L1Z5');
      expect(acme).toEqual(['29AABCT1332L1Z5', 'Acme Traders', 15000, 2700, 17700]);
      const beta = table.find((r) => r[0] === '07AAACB2230P1Z1');
      expect(beta).toEqual(['07AAACB2230P1Z1', 'Beta Supplies', 20000, 3600, 23600]);
    });

    it('does not run "first" through the numeric aggregator (no NaN, no coercion)', () => {
      const table = buildAggregateTable({
        rows: gstRows,
        hasHeaders: true,
        groupByColumn: 'GSTIN',
        aggregations: [{ column: 'Supplier Name', fn: 'first', outputLabel: 'Supplier Name' }],
      });
      const acme = table.find((r) => r[0] === '29AABCT1332L1Z5');
      expect(typeof acme?.[1]).toBe('string');
      expect(acme?.[1]).toBe('Acme Traders');
    });
  });
});
