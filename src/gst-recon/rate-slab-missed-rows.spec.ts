import { GstReconService } from './gst-recon.service';

/**
 * Regression coverage for the Ocean Polymers rate-slab bug: a real "Purchase register"
 * export with spaced/irregular rate-slab headers ("Purchase @ 18%", "Purchase Interstate @18%",
 * "Purchase @ 28 %") going through the actual missed_books_only HTTP path
 * (GstReconService.reconcile -> gstMatch -> mapMissedBooksToSheetActions), not just the
 * parser in isolation. Header values below are copied verbatim from the customer's
 * "PR vs Missed PR.xlsx" fixture (Purchase register sheet, row 11).
 */
describe('missed_books_only — rate-slab taxable value reaches the output sheet', () => {
  const service = new GstReconService();

  const realBooksHeaders = [
    'Date',
    'Particulars',
    'GSTIN/UIN',
    'Purchase@0%',
    'Purchase@5%',
    'CGST',
    'SGST',
    'Purchase @ 18%',
    'Purchase Interstate @18%',
    'IGST',
    'Purchase Interstate@28%',
    'Purchase@12%',
    'Purchase @ 28 %',
  ];

  it('carries Ocean Polymers\' taxable value (33495.78) through to the missed-rows sheet, not 0', async () => {
    const oceanPolymersRow = [
      45404,
      'Ocean Polymers',
      '32AAACF3314R1ZO',
      '',
      '',
      3014.62,
      3014.62,
      33495.78,
      '',
      '',
      '',
      '',
      '',
    ];

    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase register',
        data: [realBooksHeaders, oceanPolymersRow],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'B2B',
        data: [
          ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST', 'IGST'],
          ['27ZZZZZ0000Z1Z5', 'UNRELATED-1', '2026-04-01', 1000, 90, 90, 0],
        ],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    expect(result.summary.pr_only).toBe(1);

    const writeTable = (
      result.actions as Array<{ type: string; headers?: string[]; rows?: unknown[][] }>
    ).find((a) => a.type === 'WRITE_TABLE');
    expect(writeTable).toBeDefined();

    const headerRowIndex = writeTable!.rows!.findIndex((r) => r[0] === 'GSTIN');
    const dataRow = writeTable!.rows![headerRowIndex + 1] as unknown[];
    const columns = writeTable!.rows![headerRowIndex] as string[];

    const taxableValueCol = columns.indexOf('Taxable Value');
    expect(dataRow[taxableValueCol]).toBe(33495.78);
    expect(dataRow[taxableValueCol]).not.toBe(0);
    expect(dataRow[taxableValueCol]).not.toBeNull();

    expect(dataRow[columns.indexOf('GSTIN')]).toBe('32AAACF3314R1ZO');
    expect(dataRow[columns.indexOf('CGST')]).toBe(3014.62);
    expect(dataRow[columns.indexOf('SGST')]).toBe(3014.62);
  });
});
