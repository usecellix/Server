import { GstReconService } from './gst-recon.service';

/**
 * Follow-up "change the layout of Missed vs X" path: `relayout()` re-maps an already
 * completed casual run's rows into a different sheet layout, using the
 * `layout_cache_id` returned alongside the original result — no workbook re-read, no
 * re-matching, just the formatting step run again against the cached rows.
 */
describe('GstReconService.relayout — re-map a completed casual run without re-matching', () => {
  const service = new GstReconService();

  const booksHeaders = ['GSTIN/UIN', 'Particulars', 'Date', 'Taxable Value', 'CGST', 'SGST'];
  const booksRow = ['29ABCDE1234F1Z5', 'Acme Co', '2026-04-01', 1000, 90, 90];
  const portalHeaders = ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];
  const portalRow = ['27ZZZZZ0000Z1Z5', 'UNRELATED-1', '2026-04-01', 5000, 450, 450];

  // `reconcile()`'s return type is a union across its three sub-flows (invoice-level /
  // 3B-vs-2B); `reconciliation_type: 'PR_VS_GSTR2B'` always takes the invoice-level
  // branch, the only one carrying `layout_cache_id` — narrow the test helper's return
  // type to that branch rather than widening the service's real public type for it.
  async function reconcile(
    layout?: 'categorized' | 'books_flat' | 'portal_flat',
  ): Promise<{ output_sheet_name: string; layout_cache_id: string | null }> {
    return service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      layout,
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [booksHeaders, booksRow],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'B2B',
        data: [portalHeaders, portalRow],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any) as any;
  }

  it('a completed casual run returns a layout_cache_id', async () => {
    const result = await reconcile('categorized');
    expect(result.layout_cache_id).toBeTruthy();
  });

  it('relayout to books_flat produces the books_flat sheet name and a non-empty single-table action set', async () => {
    const original = await reconcile('categorized');
    const relayouted = service.relayout({
      layout_cache_id: original.layout_cache_id as string,
      layout: 'books_flat',
    });

    expect(relayouted.output_sheet_name).toBe('Missed vs GSTR-2B — Books Only');
    expect(relayouted.layout).toBe('books_flat');
    const createSheet = relayouted.actions.find((a) => a.type === 'CREATE_SHEET');
    expect(createSheet).toBeTruthy();
    const writeTable = relayouted.actions.find((a) => a.type === 'WRITE_TABLE');
    expect(writeTable).toBeTruthy();
  });

  it('relayout to portal_flat produces the portal_flat sheet name', async () => {
    const original = await reconcile('categorized');
    const relayouted = service.relayout({
      layout_cache_id: original.layout_cache_id as string,
      layout: 'portal_flat',
    });

    expect(relayouted.output_sheet_name).toBe('Missed vs GSTR-2B — Portal Only');
    expect(relayouted.layout).toBe('portal_flat');
  });

  it('relayout back to categorized produces the unsuffixed sheet name', async () => {
    const original = await reconcile('books_flat');
    const relayouted = service.relayout({
      layout_cache_id: original.layout_cache_id as string,
      layout: 'categorized',
    });

    expect(relayouted.output_sheet_name).toBe('Missed vs GSTR-2B');
  });

  it('an explicit output_sheet_name on the relayout request always wins', async () => {
    const original = await reconcile('categorized');
    const relayouted = service.relayout({
      layout_cache_id: original.layout_cache_id as string,
      layout: 'books_flat',
      output_sheet_name: 'My Custom Layout Sheet',
    });

    expect(relayouted.output_sheet_name).toBe('My Custom Layout Sheet');
  });

  it('an unknown or expired layout_cache_id throws rather than silently producing an empty sheet', () => {
    expect(() =>
      service.relayout({ layout_cache_id: 'layout_cache_does_not_exist', layout: 'books_flat' }),
    ).toThrow(/no longer available/i);
  });

  it('the full/gated (non-casual) reconcile path never produces a layout_cache_id — nothing to re-layout there', async () => {
    const result = (await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      client_gstin: '29ABCDE1234F1Z5',
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [booksHeaders, booksRow],
        headers_row: 1,
      },
      // The full/gated flow validates the portal file's own stated GSTIN matches
      // client_gstin before matching runs — use the same client GSTIN here so this
      // test exercises the layout_cache_id absence, not that unrelated guard.
      portal_file: {
        sheet_name: 'B2B',
        data: [portalHeaders, ['29ABCDE1234F1Z5', ...portalRow.slice(1)]],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any)) as { layout_cache_id: string | null };

    expect(result.layout_cache_id).toBeNull();
  });
});
