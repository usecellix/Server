import { GstReconService } from './gst-recon.service';

/**
 * Verifies the dual-portal-source path (both GSTR-2B and GSTR-2A resolved and sent
 * together — the casual-prompt addendum's `useDualPurchasePortals` branch) against
 * two concerns raised in sheet_name_illegal_char_fix_prompt.md:
 *
 * 1. The generated sheet name must never contain an Excel-illegal character (the
 *    "Missed vs GSTR-2B / GSTR-2A" bug — "/" is illegal in a worksheet name and
 *    crashed ADD_SHEET, cascading into every later action in the same batch).
 *
 * 2. Whether Portal-only rows are still computed/included when both portals
 *    resolve. Server-side, `portal_file_2a` rows are simply concatenated onto
 *    `portal_file` rows before matching (gst-recon.service.ts), so PORTAL_ONLY
 *    computation is mechanically source-agnostic — these tests prove that directly
 *    rather than trusting the trace. (Separately, the casual/categorized CHAT TEXT
 *    never surfaces a "Portal-only" line for ANY source combination — that's an
 *    existing, deliberate design choice of buildGstReconAnswerText, not something
 *    this suite asserts against.)
 */
describe('dual-portal-source (GSTR-2B + GSTR-2A both resolved)', () => {
  const service = new GstReconService();

  const booksHeaders = ['GSTIN/UIN', 'Particulars', 'Date', 'Taxable Value', 'CGST', 'SGST'];
  const portalHeaders = ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];

  const booksRows = [
    ['29AAAAA0000A1Z5', 'Vendor A', '2026-04-01', 1000, 90, 90], // matches a 2B row
    ['29BBBBB0000B1Z5', 'Vendor B', '2026-04-02', 2000, 180, 180], // matches a 2A row
    ['29CCCCC0000C1Z5', 'Vendor C', '2026-04-03', 3000, 270, 270], // genuinely missing
  ];

  async function reconcileDualSource(overrides: Record<string, unknown> = {}) {
    return service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [booksHeaders, ...booksRows],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'GSTR-2B',
        data: [
          portalHeaders,
          ['29AAAAA0000A1Z5', 'PB-1', '2026-04-01', 1000, 90, 90], // matches books row A
          ['27PPPPP0000P1Z5', 'INV-2B-ONLY', '2026-04-05', 5000, 450, 450], // portal-only, from 2B
        ],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      portal_file_2a: {
        sheet_name: 'GSTR-2A',
        data: [
          portalHeaders,
          ['29BBBBB0000B1Z5', 'PA-1', '2026-04-02', 2000, 180, 180], // matches books row B
          ['27QQQQQ0000Q1Z5', 'INV-2A-ONLY', '2026-04-06', 6000, 540, 540], // portal-only, from 2A
        ],
        headers_row: 1,
        file_type: 'GSTR2A',
      },
      ...overrides,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);
  }

  it('never generates a sheet name containing an Excel-illegal character', async () => {
    const result = await reconcileDualSource();
    expect(result.output_sheet_name).not.toMatch(/[\\/?*[\]:]/);
    expect(result.output_sheet_name).toBe('Missed vs GSTR-2B GSTR-2A');
  });

  it('does not crash ADD_SHEET/downstream actions — CREATE_SHEET and every other action share the identical, valid sheet name', async () => {
    const result = await reconcileDualSource();
    const sheetNames = new Set(
      (result.actions as Array<{ sheetName?: string; name?: string }>)
        .map((a) => a.sheetName ?? a.name)
        .filter(Boolean),
    );
    expect(sheetNames.size).toBe(1);
    expect([...sheetNames][0]).toBe(result.output_sheet_name);
    expect([...sheetNames][0]).not.toMatch(/[\\/?*[\]:]/);
  });

  it('computes Portal-only from BOTH sources (union) in the summary, but never writes it into the one-directional casual sheet', async () => {
    const result = await reconcileDualSource();

    expect(result.summary.pr_only).toBe(1); // Vendor C only
    // The summary count itself stays source-agnostic and correct — used for the
    // programmatic ReconSummaryBlock, not for deciding what gets written into the sheet.
    expect(result.summary.portal_only).toBe(2); // INV-2B-ONLY + INV-2A-ONLY

    const writeTable = (
      result.actions as Array<{ type: string; rows?: unknown[][] }>
    ).find((a) => a.type === 'WRITE_TABLE');
    const joined = JSON.stringify(writeTable?.rows);
    // The casual/categorized sheet is one-directional (books reconciled against portal,
    // never the reverse) — a portal-only invoice (routinely another client's, sharing the
    // same portal export) must never appear in it, from either source.
    expect(joined).not.toContain('INV-2B-ONLY');
    expect(joined).not.toContain('INV-2A-ONLY');
    expect(joined).not.toContain('Portal-only rows');
  });

  it('a long portalLabel does not truncate the layout suffix away (books_flat layout)', async () => {
    const result = await reconcileDualSource({ layout: 'books_flat' });
    expect(result.output_sheet_name.endsWith('— Books Only')).toBe(true);
    expect(result.output_sheet_name).not.toMatch(/[\\/?*[\]:]/);
  });
});

/**
 * portal_only_inflation_fix_prompt.md — the real bug: an invoice present in BOTH
 * GSTR-2B and GSTR-2A reaches the matching passes as two separate rows. Pass 1b
 * consumes one copy against the books row; the un-consumed twin used to be left
 * in `unmatchedPortal` and reported as a false-positive Portal-only row — even
 * though the invoice genuinely matched. An invoice in both sources with NO books
 * match at all was likewise double-counted in Portal-only instead of once.
 *
 * Reproduces the exact reported shape: 16 invoices matched (present, identically,
 * in both 2B and 2A), 2 invoices present in only one source each (genuinely
 * portal-only), and 1 invoice present in both sources with no books match at all
 * (also genuinely portal-only, but must be counted ONCE, not twice).
 */
describe('dual-portal-source — Portal-only deduplication (portal_only_inflation_fix_prompt.md)', () => {
  const service = new GstReconService();

  const booksHeaders = ['GSTIN/UIN', 'Particulars', 'Date', 'Taxable Value', 'CGST', 'SGST'];
  const portalHeaders = ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];

  const DUPLICATED_MATCHED_COUNT = 16;

  function buildFixture() {
    const booksRows: unknown[][] = [];
    const portal2b: unknown[][] = [];
    const portal2a: unknown[][] = [];

    // 16 invoices genuinely matched, listed identically (same GSTIN + Invoice
    // Number + Date + Amount) in BOTH portal sources.
    for (let i = 0; i < DUPLICATED_MATCHED_COUNT; i++) {
      const gstin = `29AAA${String(i).padStart(2, '0')}0000A1Z${i % 10}`;
      const invoiceNo = `DUP-${i}`;
      const amount = 1000 + i;
      const tax = Math.round(amount * 0.09 * 100) / 100;
      booksRows.push([gstin, `Vendor ${i}`, '2026-04-10', amount, tax, tax]);
      portal2b.push([gstin, invoiceNo, '2026-04-10', amount, tax, tax]);
      portal2a.push([gstin, invoiceNo, '2026-04-10', amount, tax, tax]);
    }

    // 1 invoice present in only 2B — genuinely portal-only.
    portal2b.push(['27PPPPP0000P1Z5', 'INV-2B-ONLY', '2026-04-05', 5000, 450, 450]);
    // 1 invoice present in only 2A — genuinely portal-only.
    portal2a.push(['27QQQQQ0000Q1Z5', 'INV-2A-ONLY', '2026-04-06', 6000, 540, 540]);
    // 1 invoice present in BOTH 2B and 2A, with no books match — genuinely
    // portal-only, but must be counted ONCE, not twice.
    portal2b.push(['27RRRRR0000R1Z5', 'INV-BOTH-UNMATCHED', '2026-04-07', 7000, 630, 630]);
    portal2a.push(['27RRRRR0000R1Z5', 'INV-BOTH-UNMATCHED', '2026-04-07', 7000, 630, 630]);

    return { booksRows, portal2b, portal2a };
  }

  it('Portal-only count ends up at exactly 3 genuinely-unmatched invoices — no duplicates, no already-matched invoices leaking in — and none of it is written into the one-directional casual sheet', async () => {
    const { booksRows, portal2b, portal2a } = buildFixture();

    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [booksHeaders, ...booksRows],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'GSTR-2B',
        data: [portalHeaders, ...portal2b],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      portal_file_2a: {
        sheet_name: 'GSTR-2A',
        data: [portalHeaders, ...portal2a],
        headers_row: 1,
        file_type: 'GSTR2A',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    // All 16 duplicated invoices matched; nothing left unaccounted for on the books side.
    expect(result.summary.pr_only).toBe(0);

    // Exactly 3 — not 24, not 16+3, not 6 (the two single-source ones double-counted
    // plus the both-source one double-counted). This count is still computed correctly
    // (dedupePortalRows is source-agnostic and unaffected by the sheet-content change
    // below); it's simply no longer written into the casual sheet as its own section.
    expect(result.summary.portal_only).toBe(3);

    // Nothing missed on the books side and no GSTIN mismatches — the one-directional
    // casual sheet has nothing to report, so no CREATE_SHEET/WRITE_TABLE action exists
    // at all (portal-only rows alone never trigger a sheet).
    expect(result.actions).toHaveLength(0);
  });

  it('invariant: no invoice (by GSTIN) appears in both the matched set and the Portal-only set simultaneously', async () => {
    const { booksRows, portal2b, portal2a } = buildFixture();

    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [booksHeaders, ...booksRows],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'GSTR-2B',
        data: [portalHeaders, ...portal2b],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      portal_file_2a: {
        sheet_name: 'GSTR-2A',
        data: [portalHeaders, ...portal2a],
        headers_row: 1,
        file_type: 'GSTR2A',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    const matchedGstins = new Set(
      result.rows.filter((r) => r.status === 'MATCHED').map((r) => r.gstin),
    );
    const portalOnlyGstins = result.rows
      .filter((r) => r.status === 'PORTAL_ONLY')
      .map((r) => r.gstin);

    expect(portalOnlyGstins.length).toBeGreaterThan(0);
    for (const gstin of portalOnlyGstins) {
      expect(matchedGstins.has(gstin)).toBe(false);
    }
  });
});

/**
 * BUGFIX_portal_only_double_count.md — end-to-end (through the full categorized
 * sheet output) reproduction of the exact real repro: a books row matched to a
 * portal invoice by GSTIN + amount, with a differing date, correctly lands under
 * "Date mismatch rows" — but that same portal invoice was ALSO being dumped into
 * "Portal-only rows" as if it had no books counterpart, inflating the count by
 * exactly the invoices already cited by a date_mismatch (or amount_mismatch)
 * diagnosis. Reproduced here through GstReconService.reconcile() + the categorized
 * mapper, not just the matching engine in isolation, since the bug's own report
 * was against the rendered sheet.
 */
describe('Date/amount-mismatch counterpart invoices never double-count into Portal-only (BUGFIX_portal_only_double_count.md)', () => {
  const service = new GstReconService();

  const booksHeaders = ['GSTIN/UIN', 'Particulars', 'Date', 'Taxable Value', 'CGST', 'SGST'];
  const portalHeaders = ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value', 'CGST', 'SGST'];

  it('the exact real repro: Palm Grove Textiles (PGT/312) appears once under Date mismatch, never also under Portal-only', async () => {
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: 'Purchase Register',
        data: [
          booksHeaders,
          ['32PALGT2345F1Z1', 'Palm Grove Textiles', '2024-05-25', 29000, 2610, 2610],
        ],
        headers_row: 1,
      },
      portal_file: {
        sheet_name: 'GSTR-2B',
        data: [
          portalHeaders,
          ['32PALGT2345F1Z1', 'PGT/312', '2024-05-28', 29000, 2610, 2610],
        ],
        headers_row: 1,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    expect(result.summary.mismatch_date).toBe(1);
    expect(result.summary.portal_only).toBe(0);
    expect(result.summary.pr_only).toBe(1);

    const writeTable = (result.actions as Array<{ type: string; rows?: unknown[][] }>).find(
      (a) => a.type === 'WRITE_TABLE',
    );
    const rows = writeTable?.rows ?? [];
    const joined = JSON.stringify(rows);

    // The books row (its only listing) shows up once, under Date mismatch, with
    // an explanation naming both dates — no separate "Portal-only rows" section
    // exists at all (pushSection is a no-op for zero rows), so PGT/312 — the
    // PORTAL side's own invoice number, never printed by the books-sourced Date
    // mismatch row — cannot leak in from there either.
    expect(joined).toContain('Date mismatch rows');
    expect(joined).toContain('Palm Grove Textiles');
    expect(joined).toContain('2024-05-25');
    expect(joined).toContain('2024-05-28');
    expect(joined).not.toContain('Portal-only rows');
    expect(joined).not.toContain('PGT/312');

    const occurrences = rows.filter((r) => r.includes('Palm Grove Textiles')).length;
    expect(occurrences).toBe(1);
  });
});
