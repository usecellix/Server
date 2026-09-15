import { GstReconService } from './gst-recon.service';

/**
 * Regression coverage for the Ocean Polymers Pass 1b bug: a real "Purchase register" export
 * with no Invoice Number column at all going through the actual missed_books_only HTTP path
 * (GstReconService.reconcile -> hasInvoiceNumberColumn -> gstMatch -> Pass 1b), not just
 * gstMatch() called directly. Header/row values below are copied verbatim from the customer's
 * "PR vs Missed PR.xlsx" fixture (Purchase register row 11 / B2B row 2).
 */
describe('missed_books_only — Pass 1b fallback reaches the real service path', () => {
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

  const realPortalHeaders = [
    'GSTIN',
    'NAME',
    'Invoice number',
    'Invoice type',
    'Invoice Date',
    'Invoice Value(₹)',
    'STATE',
    'RCM',
    'Taxable Value',
    'Integrated tax(₹)',
    'Central tax(₹)',
    'State/UT tax(₹)',
    'Cess(₹)',
  ];

  it('matches Ocean Polymers via Pass 1b (GSTIN+date+amount) — not PR_ONLY — when books has no Invoice Number column', async () => {
    const oceanPolymersBooksRow = [
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
    const oceanPolymersPortalRow = [
      '32AAACF3314R1ZO',
      'OCEAN POLYMER TECHNOLOGIES PRIVATE LIMITED',
      'OTH/B/573/24-25',
      'R',
      '22/04/2024',
      39525,
      'Kerala',
      'N',
      33495.78,
      0,
      3014.62,
      3014.62,
      0,
    ];

    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase register',
        data: [realBooksHeaders, oceanPolymersBooksRow],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'B2B',
        data: [realPortalHeaders, oceanPolymersPortalRow],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    // Ocean Polymers matched -> zero missed rows -> no sheet actions at all.
    expect(result.summary.pr_only).toBe(0);
    expect(result.summary.exact_matched).toBe(1);
    expect(result.actions).toEqual([]);
  });

  it('falls back to PR_ONLY (not silently 0-matched) when no plausible portal candidate exists', async () => {
    const unmatchableRow = [45404, 'Nobody Ltd', '32ZZZZZ0000Z1ZZ', '', '', 1, 1, 500, '', '', '', '', ''];

    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase register',
        data: [realBooksHeaders, unmatchableRow],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'B2B',
        data: [
          realPortalHeaders,
          ['32AAACF3314R1ZO', 'OCEAN POLYMER TECHNOLOGIES PRIVATE LIMITED', 'OTH/B/573/24-25', 'R', '22/04/2024', 39525, 'Kerala', 'N', 33495.78, 0, 3014.62, 3014.62, 0],
        ],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    expect(result.summary.pr_only).toBe(1);
    expect(result.actions.length).toBeGreaterThan(0);
  });
});
