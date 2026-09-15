import * as fs from 'fs';
import * as path from 'path';
import { GstReconService } from './gst-recon.service';
import { GstReconOrchestrator } from './orchestrator';
import { resolveSheetsForRecon } from './sheet-detector';
import { resolveMissingContext } from './context-resolver';

function loadFixture(name: string) {
  const p = path.join(__dirname, '__fixtures__', name);
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}

describe('gst-recon orchestrator + fixtures', () => {
  const service = new GstReconService();
  const orchestrator = new GstReconOrchestrator(service);

  it('needs_input when GSTIN/period/client missing', async () => {
    const result = await orchestrator.runConversationalRecon(
      'GST_RECON_PURCHASE',
      { intent: 'GST_RECON_PURCHASE' },
      [],
    );
    expect(result.kind).toBe('needs_input');
    if (result.kind === 'needs_input') {
      expect(result.prompts.length).toBeGreaterThanOrEqual(2);
    }
  });

  it('context-resolver prompts for missing fields', () => {
    const missing = resolveMissingContext({});
    expect(missing.map((m) => m.field)).toEqual(
      expect.arrayContaining(['clientGstin', 'taxPeriod', 'client']),
    );
  });

  it('ambiguous sheets fixture → chat_reply', async () => {
    const fx = loadFixture('ambiguous-sheets.json');
    const result = await orchestrator.runConversationalRecon(
      'GST_RECON_PURCHASE',
      {
        intent: 'GST_RECON_PURCHASE',
        extractedGstin: '27ABCDE1234F1Z5',
        extractedPeriod: 'April 2026',
        extractedClientName: 'ABC Traders',
      },
      fx.sheets,
    );
    expect(result.kind).toBe('chat_reply');
    if (result.kind === 'chat_reply') {
      expect(result.message).toMatch(/more than one/i);
    }
  });

  it('single-gstin-purchase → exact matches via reconcile service', async () => {
    const fx = loadFixture('single-gstin-purchase.json');
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      client_gstin: fx.clientGstin,
      period: fx.period,
      client_name: 'ABC Traders',
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: fx.booksGrid,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR2B',
      },
    });
    expect(result.summary.exact_matched).toBeGreaterThanOrEqual(2);
    expect(
      (result as { cross_gstin_exception_count?: number }).cross_gstin_exception_count ?? 0,
    ).toBe(0);
  });

  it('single-gstin-purchase conversational path returns action_payload', async () => {
    const fx = loadFixture('single-gstin-purchase.json');
    const sheets = [
      { sheetName: fx.booksSheet, headers: fx.booksHeaders },
      { sheetName: fx.portalSheet, headers: fx.portalHeaders },
    ];
    const result = await orchestrator.runConversationalRecon(
      'GST_RECON_PURCHASE',
      {
        intent: 'GST_RECON_PURCHASE',
        extractedGstin: fx.clientGstin,
        extractedPeriod: fx.period,
        extractedClientName: 'ABC Traders',
      },
      sheets,
      {
        [fx.booksSheet]: fx.booksGrid,
        [fx.portalSheet]: fx.portalGrid,
      },
    );
    expect(result.kind).toBe('action_payload');
  });

  it('portal-gstin-mismatch → error hard stop', async () => {
    const fx = loadFixture('portal-gstin-mismatch.json');
    const result = await orchestrator.runConversationalRecon(
      'GST_RECON_PURCHASE',
      {
        intent: 'GST_RECON_PURCHASE',
        extractedGstin: fx.clientGstin,
        extractedPeriod: fx.period,
        extractedClientName: 'ABC Traders',
      },
      [
        {
          sheetName: fx.booksSheet,
          headers: ['Supplier GSTIN', 'Invoice No', 'Invoice Date', 'Taxable Value'],
        },
        {
          sheetName: fx.portalSheet,
          headers: [
            'GSTIN of supplier',
            'Invoice number',
            'Invoice Date',
            'Document Type',
            'ITC Available',
          ],
        },
      ],
      {
        [fx.booksSheet]: fx.booksGrid,
        [fx.portalSheet]: fx.portalGrid,
      },
    );
    expect(result.kind).toBe('error');
    if (result.kind === 'error') {
      expect(result.message).toMatch(/does not match/i);
    }
  });

  it('multi-branch-cross-gstin flags exceptions', async () => {
    const fx = loadFixture('multi-branch-cross-gstin.json');
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      client_gstin: fx.clientGstin,
      period: fx.period,
      client_name: 'ABC Traders',
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: fx.booksGrid,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR2B',
      },
    });
    expect((result as { cross_gstin_exception_count?: number }).cross_gstin_exception_count).toBeGreaterThanOrEqual(1);
    expect(result.rows.some((r: { status: string }) => r.status === 'CROSS_GSTIN')).toBe(true);
  });

  it('sales B2B+B2C mixed matches without invalid_missing_gstin on B2C', async () => {
    const fx = loadFixture('single-gstin-sales-b2b-b2c-mixed.json');
    const result = await service.reconcile({
      reconciliation_type: 'SALES_VS_GSTR1',
      client_gstin: fx.clientGstin,
      period: fx.period,
      client_name: 'XYZ Pvt Ltd',
      books_register: {
        sheet_name: fx.booksSheet,
        data: fx.booksGrid,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR1',
      },
    });
    expect(result.summary.exact_matched).toBeGreaterThanOrEqual(2);
    expect(result.reconciliation_type).toBe('SALES_VS_GSTR1');
  });

  it('missed_books_only without GSTIN creates a missed-rows sheet', async () => {
    const fx = loadFixture('single-gstin-purchase.json');
    const books = [
      ...fx.booksGrid,
      ['27CCCCC0000C1Z5', 'INV-003', '2026-04-03', 2000, 180, 180, 0],
    ];
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: books,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR2B',
      },
    });
    expect((result as { missed_books_only?: boolean }).missed_books_only).toBe(true);
    expect(result.summary.pr_only).toBeGreaterThanOrEqual(1);
    expect(result.actions.some((a: { type: string }) => a.type === 'CREATE_SHEET')).toBe(
      true,
    );
  });

  it('missed_books_only with no misses returns no sheet actions', async () => {
    const fx = loadFixture('single-gstin-purchase.json');
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: fx.booksGrid,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR2B',
      },
    });
    expect(result.summary.pr_only).toBe(0);
    expect(result.actions).toEqual([]);
  });

  it('missed_books_only merges GSTR-2A so a 2A-only invoice is not missed', async () => {
    const fx = loadFixture('single-gstin-purchase.json');
    const books = [
      ...fx.booksGrid,
      ['27CCCCC0000C1Z5', 'INV-2A', '2026-04-04', 3000, 270, 270, 0],
    ];
    const gstr2a = [
      ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST', 'IGST'],
      ['27CCCCC0000C1Z5', 'INV-2A', '2026-04-04', 3000, 270, 270, 0],
    ];
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: books,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow ?? 2,
        file_type: 'GSTR2B',
      },
      portal_file_2a: {
        sheet_name: 'GSTR-2A',
        data: gstr2a,
        headers_row: 1,
        file_type: 'GSTR2A',
      },
    });
    expect((result as { portal_label?: string }).portal_label).toMatch(/2A/);
    expect(result.rows.some((r: { invoice_number?: string | null; status: string }) => r.invoice_number === 'INV-2A' && r.status === 'PR_ONLY')).toBe(
      false,
    );
  });

  it('resolveSheetsForRecon on ambiguous fixture', () => {
    const fx = loadFixture('ambiguous-sheets.json');
    const resolved = resolveSheetsForRecon(fx.sheets, [
      'purchase_register',
      'gstr_2b',
    ]);
    expect(resolved.purchase_register.status).toBe('ambiguous');
    expect(resolved.gstr_2b.status).toBe('resolved');
  });
});
