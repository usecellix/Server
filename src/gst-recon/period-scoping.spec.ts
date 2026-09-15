import { GstReconService } from './gst-recon.service';

/**
 * Regression coverage for period scoping: a stated period (e.g. "April 2024") must filter
 * BOTH books and portal rows to that date range BEFORE matching runs — not just label the
 * output afterwards. Previously the period was extracted from the prompt but never applied,
 * so a period-scoped request silently reconciled the entire file.
 */
describe('period scoping — filters books and portal rows before matching', () => {
  const service = new GstReconService();

  const headers = ['GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];
  // 2 rows in April 2024, 2 rows in a different month, for both books and portal.
  const booksGrid = [
    headers,
    ['27AAAAA0000A1Z5', 'APR-1', '05/04/2024', 1000, 90, 90],
    ['27BBBBB0000B1Z5', 'APR-2', '20/04/2024', 2000, 180, 180],
    ['27CCCCC0000C1Z5', 'MAY-1', '05/05/2024', 3000, 270, 270],
    ['27DDDDD0000D1Z5', 'MAY-2', '20/05/2024', 4000, 360, 360],
  ];
  const portalGrid = [
    ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST', 'IGST'],
    ['27AAAAA0000A1Z5', 'APR-1', '05/04/2024', 1000, 90, 90, 0],
    ['27BBBBB0000B1Z5', 'APR-2', '20/04/2024', 2000, 180, 180, 0],
    ['27CCCCC0000C1Z5', 'MAY-1', '05/05/2024', 3000, 270, 270, 0],
    ['27DDDDD0000D1Z5', 'MAY-2', '20/05/2024', 4000, 360, 360, 0],
  ];

  const base = {
    reconciliation_type: 'PR_VS_GSTR2B' as const,
    missed_books_only: true,
    purchase_register: { sheet_name: 'Purchase register', data: booksGrid, headers_row: 1 },
    portal_file: { sheet_name: 'B2B', data: portalGrid, headers_row: 1, file_type: 'GSTR2B' },
  };

  it('with no period stated, processes every row (unchanged behavior)', async () => {
    const result = await service.reconcile(base as never);
    expect(result.summary.total_pr_rows).toBe(4);
    expect(result.summary.total_portal_rows).toBe(4);
    expect(result.summary.exact_matched).toBe(4);
  });

  it('with a stated period, filters both books and portal to that range before matching', async () => {
    const result = await service.reconcile({
      ...base,
      period_start: '2024-04-01',
      period_end: '2024-04-30',
      period_label: 'April 2024',
    } as never);

    expect(result.summary.total_pr_rows).toBe(2);
    expect(result.summary.total_portal_rows).toBe(2);
    expect(result.summary.exact_matched).toBe(2);
    expect((result as { period_applied?: { label: string } }).period_applied).toEqual({
      label: 'April 2024',
      start: '2024-04-01',
      end: '2024-04-30',
    });
  });

  it('a period matching zero books rows returns a chat message instead of running an empty reconciliation', async () => {
    const result = await service.reconcile({
      ...base,
      period_start: '2024-01-01',
      period_end: '2024-01-31',
      period_label: 'January 2024',
    } as never);

    const zeroMsg = (result as { period_zero_message?: string }).period_zero_message;
    expect(zeroMsg).toBeDefined();
    expect(zeroMsg).toContain('January 2024');
    expect(zeroMsg).toMatch(/no purchase register rows found/i);
    expect(result.actions).toEqual([]);
    expect(result.summary.exact_matched).toBe(0);
  });

  it('a period matching books but zero portal rows also short-circuits with a message', async () => {
    const bookOnlyAprilGrid = [
      headers,
      ['27ZZZZZ0000Z1Z5', 'APR-ONLY', '10/04/2024', 500, 45, 45],
    ];
    const emptyAprilPortalGrid = [
      ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST', 'IGST'],
      ['27CCCCC0000C1Z5', 'MAY-1', '05/05/2024', 3000, 270, 270, 0],
    ];
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: { sheet_name: 'Purchase register', data: bookOnlyAprilGrid, headers_row: 1 },
      portal_file: { sheet_name: 'B2B', data: emptyAprilPortalGrid, headers_row: 1, file_type: 'GSTR2B' },
      period_start: '2024-04-01',
      period_end: '2024-04-30',
      period_label: 'April 2024',
    } as never);

    const zeroMsg = (result as { period_zero_message?: string }).period_zero_message;
    expect(zeroMsg).toBeDefined();
    expect(zeroMsg).toMatch(/no gstr-2b rows found/i);
    expect(result.actions).toEqual([]);
  });
});
