import { parsePurchaseRegister } from './purchase-register-parser';
import { parseSalesRegister } from './sales-register-parser';

describe('parsePurchaseRegister — rate-slab column layout', () => {
  const headers = ['GSTIN', 'Invoice No', 'Invoice Date', 'Purchase@18%', 'Purchase Interstate@18%'];

  it('derives taxable value from a populated standard-rate column', () => {
    const grid = [
      headers,
      ['29AAAAA0000A1Z5', 'INV-1', '2026-04-01', 10000, ''],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBe(10000);
    expect(rows[0].taxRatePercent).toBe(18);
  });

  it('derives taxable value from a populated interstate column', () => {
    const grid = [
      headers,
      ['29AAAAA0000A1Z5', 'INV-2', '2026-04-02', '', 5000],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBe(5000);
    expect(rows[0].taxRatePercent).toBe(18);
  });

  it('leaves taxable value null (never 0) when every slab column is blank', () => {
    const grid = [
      headers,
      ['29AAAAA0000A1Z5', 'INV-3', '2026-04-03', '', ''],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBeNull();
  });

  it('does not treat a standard single Taxable Value layout as rate-slab', () => {
    const grid = [
      ['GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value'],
      ['29AAAAA0000A1Z5', 'INV-4', '2026-04-04', 7000],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows[0].taxableValue).toBe(7000);
    expect(rows[0].taxRatePercent).toBeUndefined();
  });

  it('regression: Ocean Polymers row with real-world spaced headers ("Purchase @ 18%") is not dropped to 0/null', () => {
    const spacedHeaders = [
      'GSTIN',
      'Invoice No',
      'Invoice Date',
      'Purchase @ 0%',
      'Purchase @ 5%',
      'Purchase @ 12%',
      'Purchase @ 18%',
      'Purchase @ 28 %',
      'Purchase Interstate @18%',
      'CGST',
      'SGST',
    ];
    const grid = [
      spacedHeaders,
      [
        '32AAACF3314R1ZO',
        'OP-001',
        '22/04/2024',
        '',
        '',
        '',
        33495.78,
        '',
        '',
        3014.62,
        3014.62,
      ],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].gstin).toBe('32AAACF3314R1ZO');
    expect(rows[0].taxableValue).toBe(33495.78);
    expect(rows[0].taxRatePercent).toBe(18);
    expect(rows[0].cgst).toBe(3014.62);
    expect(rows[0].sgst).toBe(3014.62);
  });

  const realHeaders = [
    'Date', 'Particulars', 'GSTIN/UIN', 'Purchase@0%', 'Purchase@5%', 'CGST', 'SGST',
    'Purchase @ 18%', 'Purchase Interstate @18%', 'IGST', 'Purchase Interstate@28%',
    'Purchase@12%', 'Purchase @ 28 %',
  ];

  it('regression (row 133): two non-blank slab columns — picks the one whose implied tax matches actual CGST+SGST, not the first in priority order', () => {
    // Real row: Purchase@12%=141.96 AND Purchase@18%=17166.11, CGST+SGST=3106.94 (= 18% of 17166.11).
    // Purchase@12% sorts before Purchase@18% in priority order — the bug picked 141.96.
    const grid = [
      realHeaders,
      ['01/04/2026', 'Zigma Tools', '32AADFZ8303N1ZW', '', '', 1553.47, 1553.47, 17166.11, '', '', '', 141.96, ''],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBe(17166.11);
    expect(rows[0].taxRatePercent).toBe(18);
    expect(rows[0].ambiguousRateSlab).toBeUndefined();
  });

  it('regression (row 142): two non-blank slab columns with no close tax match — flags ambiguous_rate_slab, never guesses', () => {
    // Real row: Purchase@12%=4017.86 AND Purchase@18%=4555.08, CGST+SGST=1302.06 — a genuine
    // blended-rate consolidated line (819.91 @18% + 482.14 @12% ≈ 1302.06), not a single slab.
    const grid = [
      realHeaders,
      ['01/04/2026', 'Zigma Tools', '32AADFZ8303N1ZW', '', '', 651.03, 651.03, 4555.08, '', '', '', 4017.86, ''],
    ];
    const rows = parsePurchaseRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBeNull();
    expect(rows[0].ambiguousRateSlab).toBe(true);
    expect(rows[0].ambiguousRateSlabDetail).toMatch(/multiple rate-slab columns/i);
  });
});

describe('parseSalesRegister — rate-slab column layout', () => {
  const headers = ['Recipient GSTIN', 'Invoice No', 'Invoice Date', 'Sales@18%', 'Sales Interstate@18%'];

  it('derives taxable value from a populated standard-rate column', () => {
    const grid = [
      headers,
      ['29BBBBB0000B1Z5', 'SINV-1', '2026-04-01', 8000, ''],
    ];
    const rows = parseSalesRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBe(8000);
    expect(rows[0].taxRatePercent).toBe(18);
  });

  it('derives taxable value from a populated interstate column', () => {
    const grid = [
      headers,
      ['29BBBBB0000B1Z5', 'SINV-2', '2026-04-02', '', 3000],
    ];
    const rows = parseSalesRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBe(3000);
    expect(rows[0].taxRatePercent).toBe(18);
  });

  it('leaves taxable value null (never 0) when every slab column is blank', () => {
    const grid = [
      headers,
      ['29BBBBB0000B1Z5', 'SINV-3', '2026-04-03', '', ''],
    ];
    const rows = parseSalesRegister(grid);
    expect(rows).toHaveLength(1);
    expect(rows[0].taxableValue).toBeNull();
  });
});
