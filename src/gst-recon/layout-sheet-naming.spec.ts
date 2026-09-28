import { GstReconService } from './gst-recon.service';

/**
 * Each missed_books_only layout must write to its own distinct sheet name — otherwise
 * switching layouts (e.g. running the default, then asking for the books-only view)
 * would trip the Overwrite/Create-new collision flow instead of creating a sibling
 * sheet. That flow is only meant to fire when the SAME layout is re-run.
 */
describe('missed_books_only — each layout resolves to its own sheet name', () => {
  const service = new GstReconService();

  const booksHeaders = ['GSTIN/UIN', 'Particulars', 'Date', 'Taxable Value', 'CGST', 'SGST'];
  const booksRow = ['29ABCDE1234F1Z5', 'Acme Co', '2026-04-01', 1000, 90, 90];
  const portalHeaders = ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];
  const portalRow = ['27ZZZZZ0000Z1Z5', 'UNRELATED-1', '2026-04-01', 5000, 450, 450];

  async function reconcile(layout?: 'categorized' | 'books_flat' | 'portal_flat') {
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
    } as any);
  }

  it('defaults to "Missed vs GSTR-2B" when layout is unspecified — unchanged from before this feature', async () => {
    const result = await reconcile(undefined);
    expect(result.output_sheet_name).toBe('Missed vs GSTR-2B');
  });

  it('categorized explicitly resolves to the same unsuffixed name as the default', async () => {
    const result = await reconcile('categorized');
    expect(result.output_sheet_name).toBe('Missed vs GSTR-2B');
  });

  it('books_flat resolves to a distinct suffixed name', async () => {
    const result = await reconcile('books_flat');
    expect(result.output_sheet_name).toBe('Missed vs GSTR-2B — Books Only');
  });

  it('portal_flat resolves to a distinct suffixed name', async () => {
    const result = await reconcile('portal_flat');
    expect(result.output_sheet_name).toBe('Missed vs GSTR-2B — Portal Only');
  });

  it('all three layouts resolve to mutually distinct sheet names, so switching layouts never collides', async () => {
    const [categorized, booksFlat, portalFlat] = await Promise.all([
      reconcile('categorized'),
      reconcile('books_flat'),
      reconcile('portal_flat'),
    ]);
    const names = [categorized.output_sheet_name, booksFlat.output_sheet_name, portalFlat.output_sheet_name];
    expect(new Set(names).size).toBe(3);
  });

  it('an explicit output_sheet_name always wins, regardless of layout', async () => {
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      layout: 'books_flat',
      output_sheet_name: 'My Custom Sheet',
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
    } as any);
    expect(result.output_sheet_name).toBe('My Custom Sheet');
  });
});
