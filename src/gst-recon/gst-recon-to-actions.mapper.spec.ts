import {
  buildSummary,
  mapMissedBooksFlatToSheetActions,
  mapMissedBooksToSheetActions,
  mapPortalOnlyFlatToSheetActions,
  mapResultRows,
  mapToSheetActions,
} from './gst-recon-to-actions.mapper';
import { GstResultRow } from '../domain-tools/gst/match-passes';
import { NormalizedInvoiceRow } from '../domain-tools/types/domain-tool.types';

let rowCounter = 0;

function booksRow(invoice: string, overrides: Partial<NormalizedInvoiceRow> = {}): NormalizedInvoiceRow {
  rowCounter += 1;
  return {
    gstin: '27AAAAA0000A1Z5',
    invoiceNumber: invoice,
    normalizedInvoiceNumber: invoice.replace(/[^A-Z0-9]/gi, '').toUpperCase(),
    invoiceDate: '2026-04-01',
    taxableValue: 1000,
    taxAmount: 180,
    igst: 0,
    cgst: 90,
    sgst: 90,
    narration: '',
    documentType: 'invoice',
    sourceRowRef: {
      documentType: 'purchase_register',
      documentId: 'PR',
      rowOrLine: rowCounter,
    },
    ...overrides,
  };
}

function portalRow(invoice: string, overrides: Partial<NormalizedInvoiceRow> = {}): NormalizedInvoiceRow {
  return booksRow(invoice, {
    sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: rowCounter },
    ...overrides,
  });
}

function writeTableRows(actions: ReturnType<typeof mapMissedBooksToSheetActions>): unknown[][] {
  const write = actions.find((a) => a.type === 'WRITE_TABLE');
  if (!write || write.type !== 'WRITE_TABLE' || !write.rows) throw new Error('no WRITE_TABLE action');
  return write.rows;
}

describe('mapResultRows — books_sheet_name/books_row let the client jump to the exact cell', () => {
  it('populates books_sheet_name/books_row from the books row\'s real sourceRowRef, regardless of its domain documentType label', () => {
    const pr = booksRow('INV-1', {
      narration: 'Some Vendor',
      sourceRowRef: { documentType: 'purchase_register', documentId: 'Purchase Reg', rowOrLine: 9 },
    });
    const rows: GstResultRow[] = [
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: pr,
        mismatchReason: 'genuinely_missing',
      },
    ];
    const dtos = mapResultRows(rows);
    expect(dtos[0].books_sheet_name).toBe('Purchase Reg');
    expect(dtos[0].books_row).toBe(9);
  });

  it('is null for a row with no books-side counterpart (e.g. a pure portal-only row)', () => {
    const portal = portalRow('INV-2');
    const rows: GstResultRow[] = [
      {
        status: 'PORTAL_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        portalRow: portal,
      },
    ];
    const dtos = mapResultRows(rows);
    expect(dtos[0].books_sheet_name).toBeNull();
    expect(dtos[0].books_row).toBeNull();
  });
});

describe('buildSummary — matched-by-pass and unmatched-by-reason breakdown', () => {
  it('never collapses matched into one number and unmatched into one number', () => {
    const resultRows: GstResultRow[] = [
      { status: 'MATCHED', pass: 1, confidence: 1, itcAmount: 100, rcmFlag: false },
      { status: 'MATCHED', pass: 2, confidence: 0.95, itcAmount: 100, rcmFlag: false },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        mismatchReason: 'blank_counterparty_gstin',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        mismatchReason: 'amount_mismatch',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        mismatchReason: 'gstin_not_in_portal',
      },
      {
        status: 'GSTIN_MISMATCH',
        pass: null,
        confidence: 0.85,
        itcAmount: 0,
        rcmFlag: false,
        mismatchReason: 'gstin_mismatch_same_pan',
      },
    ];
    const itc = {
      totalItcClaimable: 0,
      totalItcAtRisk: 0,
      totalRcmPayable: 0,
      totalImsRejected: 0,
      totalImsPending: 0,
    } as never;
    const summary = buildSummary(resultRows, itc, 6, 2, 0);

    expect(summary.matched_exact).toBe(1);
    expect(summary.matched_fallback).toBe(1);
    expect(summary.mismatch_blank_gstin).toBe(1);
    expect(summary.mismatch_amount).toBe(1);
    expect(summary.mismatch_genuinely_missing).toBe(1);
    expect(summary.pr_only).toBe(3);
    // GSTIN_MISMATCH rows are their own bucket — never counted in pr_only or portal_only.
    expect(summary.gstin_mismatch_count).toBe(1);
  });
});

describe('mapMissedBooksToSheetActions', () => {
  it('returns no actions when every books row matched and there are no portal-only rows', () => {
    const rows: GstResultRow[] = [
      {
        status: 'MATCHED',
        pass: 1,
        confidence: 1,
        itcAmount: 180,
        rcmFlag: false,
        registerRow: booksRow('INV-001'),
      },
    ];
    expect(
      mapMissedBooksToSheetActions({
        sheetName: 'Missed vs GSTR-2B',
        portalLabel: 'GSTR-2B',
        booksLabel: 'Purchase Register',
        runAt: '2026-04-01T00:00:00.000Z',
        resultRows: rows,
      }),
    ).toEqual([]);
  });

  it('groups unmatched rows into labeled sections by mismatch reason, not one flat list', () => {
    const rows: GstResultRow[] = [
      {
        status: 'MATCHED',
        pass: 1,
        confidence: 1,
        itcAmount: 180,
        rcmFlag: false,
        registerRow: booksRow('INV-MATCHED'),
      },
      {
        status: 'MATCHED',
        pass: 2,
        confidence: 0.95,
        itcAmount: 180,
        rcmFlag: false,
        diffType: 'FALLBACK_NO_INVOICE_NUMBER',
        registerRow: booksRow('INV-FALLBACK'),
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('', { gstin: '', narration: 'Blank GSTIN Vendor' }),
        mismatchReason: 'blank_counterparty_gstin',
        explanation: 'GSTIN is blank in the register for this row — cannot be matched.',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-BLANK-VALUE', { taxableValue: null, narration: 'Blank Value Vendor' }),
        mismatchReason: 'blank_taxable_value',
        explanation: 'No taxable value found in any rate column for this row.',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-AMBIGUOUS-SLAB', {
          taxableValue: null,
          narration: 'Ambiguous Slab Vendor',
          ambiguousRateSlab: true,
          ambiguousRateSlabDetail: 'Multiple rate-slab columns are populated (...) but none matches.',
        }),
        mismatchReason: 'ambiguous_rate_slab',
        explanation: 'Multiple rate-slab columns are populated (...) but none matches.',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-AMT-MISMATCH', { taxableValue: 10000, narration: 'Amount Mismatch Vendor' }),
        mismatchReason: 'amount_mismatch',
        explanation: 'Same GSTIN and date found in portal, but amount differs (books ₹10000 vs portal ₹15000).',
        fieldDiff: [{ field: 'taxableValue', booksValue: 10000, portalValue: 15000 }],
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-DATE-MISMATCH', {
          invoiceDate: '2026-04-15',
          narration: 'Date Mismatch Vendor',
        }),
        mismatchReason: 'date_mismatch',
        explanation: 'Same GSTIN and amount found in portal, but date differs (books 2026-04-15 vs portal 2026-05-20).',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-NOT-IN-PORTAL', { narration: 'Not In Portal Vendor' }),
        mismatchReason: 'gstin_not_in_portal',
        explanation: 'GSTIN does not appear anywhere in the portal file for this period.',
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-GENUINELY-MISSING', { narration: 'Genuinely Missing Vendor' }),
        mismatchReason: 'genuinely_missing',
        explanation: 'No matching invoice found in the portal for this GSTIN, date, or amount.',
      },
      {
        // A PORTAL_ONLY row must never appear in the categorized casual-flow sheet —
        // this is a one-directional books-vs-portal reconciliation, so a row that only
        // exists on the portal side (which, in real usage, routinely includes OTHER
        // clients' invoices sharing the same portal export) is never its own category.
        status: 'PORTAL_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        portalRow: portalRow('PORTAL-ONLY-1', { narration: 'Portal Only Vendor' }),
      },
    ];

    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      relativeTo: 'PR',
      resultRows: rows,
    });

    expect(actions[0]).toMatchObject({ type: 'CREATE_SHEET', sheetName: 'Missed vs GSTR-2B' });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    // Every unmatched row appears somewhere; matched rows never appear in the detail sections.
    expect(joined).not.toContain('INV-MATCHED');
    expect(joined).not.toContain('INV-FALLBACK');
    expect(joined).toContain('INV-BLANK-VALUE');
    expect(joined).toContain('INV-AMBIGUOUS-SLAB');
    expect(joined).toContain('INV-AMT-MISMATCH');
    expect(joined).toContain('INV-DATE-MISMATCH');
    expect(joined).toContain('INV-NOT-IN-PORTAL');
    expect(joined).toContain('INV-GENUINELY-MISSING');
    // One-directional: a portal-only row is never written into this sheet.
    expect(joined).not.toContain('PORTAL-ONLY-1');
    expect(joined).not.toContain('Portal-only rows');

    // Grouped section titles are present — never a single flat unmatched table.
    const sectionTitles = rowsOut.map((r) => r[0]).filter((c) => typeof c === 'string');
    expect(sectionTitles).toEqual(
      expect.arrayContaining([
        'Blank GSTIN rows',
        'Blank taxable value rows',
        'Ambiguous rate slab rows (needs CA review)',
        'Amount mismatch rows',
        'Date mismatch rows',
        'Genuinely missing rows',
      ]),
    );

    // Summary counts break down matched-by-pass and unmatched-by-reason.
    expect(joined).toContain('Matched (exact)');
    expect(joined).toContain('Matched (fallback, no invoice no.)');
    expect(joined).toContain('Blank GSTIN in books');
    expect(joined).toContain('Blank taxable value');
    expect(joined).toContain('Ambiguous rate slab (needs CA review)');
    expect(joined).toContain('Amount mismatch');
    expect(joined).toContain('Date mismatch');
    expect(joined).toContain('Genuinely missing');

    // Every section header row has GSTIN immediately followed by Vendor Name.
    for (const title of [
      'Blank GSTIN rows',
      'Blank taxable value rows',
      'Ambiguous rate slab rows (needs CA review)',
      'Amount mismatch rows',
      'Date mismatch rows',
      'Genuinely missing rows',
    ]) {
      const idx = rowsOut.findIndex((r) => r[0] === title);
      expect(rowsOut[idx + 1].slice(0, 2)).toEqual(['GSTIN', 'Vendor Name']);
    }
    // One-directional: no portal-only section exists in this sheet at all.
    expect(rowsOut.some((r) => String(r[0]).startsWith('Portal-only rows'))).toBe(false);

    // Vendor name (from Particulars/narration) appears in column 2 of each books-side row.
    const findDataRow = (invoiceNumber: string) =>
      rowsOut.find((r) => r[2] === invoiceNumber || r[1] === invoiceNumber);
    expect(findDataRow('INV-BLANK-VALUE')?.[1]).toBe('Blank Value Vendor');
    expect(findDataRow('INV-AMBIGUOUS-SLAB')?.[1]).toBe('Ambiguous Slab Vendor');
    expect(findDataRow('INV-AMT-MISMATCH')?.[1]).toBe('Amount Mismatch Vendor');
    expect(findDataRow('INV-DATE-MISMATCH')?.[1]).toBe('Date Mismatch Vendor');
    expect(findDataRow('INV-NOT-IN-PORTAL')?.[1]).toBe('Not In Portal Vendor');
    expect(findDataRow('INV-GENUINELY-MISSING')?.[1]).toBe('Genuinely Missing Vendor');
    // The blank-GSTIN row has an empty invoice number, so find it by its vendor name instead.
    expect(joined).toContain('Blank GSTIN Vendor');

    // genuinely_missing and gstin_not_in_portal are combined under one "Genuinely missing" section (2 rows).
    const genuinelyMissingSectionIdx = rowsOut.findIndex((r) => r[0] === 'Genuinely missing rows');
    const headerRow = rowsOut[genuinelyMissingSectionIdx + 1];
    expect(headerRow[0]).toBe('GSTIN');
    const dataRows = rowsOut
      .slice(genuinelyMissingSectionIdx + 2)
      .filter((r) => typeof r[2] === 'string' && (r[2] as string).startsWith('INV-'));
    // both genuinely_missing and gstin_not_in_portal rows land in this one section
    expect(dataRows.some((r) => r[2] === 'INV-NOT-IN-PORTAL')).toBe(true);
    expect(dataRows.some((r) => r[2] === 'INV-GENUINELY-MISSING')).toBe(true);
  });

  it('writes a "Possible GSTIN mismatch" section after Genuinely missing, with both GSTINs side by side, and never a portal-only section', () => {
    const explanation =
      'Same vendor (PAN AIOPJ2231N), same date and amount, but booked under GSTIN 32AIOPJ2231N1Z8 in your register ' +
      'vs GSTIN 33AIOPJ2231N1Z6 in GSTR-2B — check which registration this vendor actually used for this invoice.';
    const rows: GstResultRow[] = [
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-GENUINELY-MISSING-2', { narration: 'Some Other Vendor' }),
        mismatchReason: 'genuinely_missing',
        explanation: 'No matching invoice found.',
      },
      {
        status: 'GSTIN_MISMATCH',
        pass: null,
        confidence: 0.85,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('', {
          gstin: '32AIOPJ2231N1Z8',
          narration: 'Deva Steels',
          taxableValue: 120780.4,
        }),
        portalRow: portalRow('CBR/24-25/1691', {
          gstin: '33AIOPJ2231N1Z6',
          narration: 'DEVA  STEELS',
          taxableValue: 120780.4,
        }),
        mismatchReason: 'gstin_mismatch_same_pan',
        explanation,
      },
      {
        // Never surfaced in this sheet — one-directional books-vs-portal only.
        status: 'PORTAL_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        portalRow: portalRow('PORTAL-ONLY-2', { narration: 'Some Portal Vendor' }),
      },
    ];

    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    const rowsOut = writeTableRows(actions);
    const titles = rowsOut.map((r) => r[0]);

    // The section title text also appears as a summary count line earlier in the sheet —
    // find the actual SECTION (title row immediately followed by its header row), not that line.
    const genuinelyMissingIdx = titles.indexOf('Genuinely missing rows');
    const gstinMismatchIdx = rowsOut.findIndex(
      (r, i) =>
        r[0] === 'Possible GSTIN mismatch (same vendor, different registration)' &&
        rowsOut[i + 1]?.[0] === 'Vendor Name',
    );

    expect(genuinelyMissingIdx).toBeGreaterThan(-1);
    expect(gstinMismatchIdx).toBeGreaterThan(genuinelyMissingIdx);
    // One-directional: a portal-only row/section is never written into this sheet, even
    // when a GSTIN-mismatch row is also present.
    expect(titles.some((t) => String(t).startsWith('Portal-only rows'))).toBe(false);
    expect(JSON.stringify(rowsOut)).not.toContain('PORTAL-ONLY-2');

    const headerRow = rowsOut[gstinMismatchIdx + 1];
    expect(headerRow.slice(0, 3)).toEqual(['Vendor Name', 'Books GSTIN', 'Portal GSTIN']);
    const dataRow = rowsOut[gstinMismatchIdx + 2];
    expect(dataRow[0]).toBe('Deva Steels');
    expect(dataRow[1]).toBe('32AIOPJ2231N1Z8');
    expect(dataRow[2]).toBe('33AIOPJ2231N1Z6');

    const joined = JSON.stringify(rowsOut);
    expect(joined).toContain('Possible GSTIN mismatch (same vendor, different registration)');
    expect(joined).toContain('CBR/24-25/1691');
    // The GSTIN-mismatch row must never also appear in the genuinely-missing or portal-only sections.
    const genuinelyMissingSectionRows = rowsOut.slice(genuinelyMissingIdx + 2, gstinMismatchIdx);
    expect(genuinelyMissingSectionRows.some((r) => r[2] === '32AIOPJ2231N1Z8')).toBe(false);
  });

  it('writes a "Possible RCM" section for RCM-status rows, with a summary count line and detail section', () => {
    const rows: GstResultRow[] = [
      {
        status: 'RCM',
        pass: 5,
        confidence: 0.7,
        itcAmount: 0,
        rcmFlag: true,
        difference: 'RCM keyword in narration: Freight charges via GTA',
        registerRow: booksRow('FRT-1', { narration: 'Freight charges via GTA', taxableValue: 8000 }),
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    expect(joined).toContain('Possible RCM (reverse charge — needs CA review)');
    expect(joined).toContain('FRT-1');
    expect(joined).toContain('Freight charges via GTA');

    const sectionIdx = rowsOut.findIndex(
      (r) => r[0] === 'Possible RCM rows (reverse charge — needs CA review)',
    );
    expect(sectionIdx).toBeGreaterThan(-1);
    expect(rowsOut[sectionIdx + 1].slice(0, 2)).toEqual(['GSTIN', 'Vendor Name']);
  });

  it('writes an "Amended invoices" section for AMENDED-status rows, showing both books and portal invoice numbers', () => {
    const books = booksRow('INV-100', { narration: 'Zenith Traders', taxableValue: 5000 });
    const portal = portalRow('IA-100', { narration: 'ZENITH TRADERS', taxableValue: 5000 });
    const rows: GstResultRow[] = [
      {
        status: 'AMENDED',
        pass: 4,
        confidence: 0.98,
        itcAmount: 900,
        rcmFlag: false,
        difference: 'Amended invoice: INV-100 ↔ IA-100 — verify against the original invoice',
        registerRow: books,
        portalRow: portal,
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    expect(joined).toContain('Amended invoices (verify against original)');
    expect(joined).toContain('INV-100');
    expect(joined).toContain('IA-100');
    expect(joined).not.toContain('CREDIT_NOTE');
  });

  it('an RCM-only or Amended-only result still produces a sheet (not an empty return)', () => {
    const rcmOnly: GstResultRow[] = [
      {
        status: 'RCM',
        pass: 5,
        confidence: 0.7,
        itcAmount: 0,
        rcmFlag: true,
        registerRow: booksRow('FRT-2', { narration: 'Freight via GTA' }),
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rcmOnly,
    });
    expect(actions.length).toBeGreaterThan(0);
  });

  it('writes a "Blank GSTIN — likely matched" section right after "Blank GSTIN rows", with the suggested GSTIN shown, and never lists the resolved portal row separately', () => {
    const explanation =
      'GSTIN is blank in the register, but a portal invoice from FABS TRADING COMPANY (GSTIN 32AAGFF3216M1ZL) ' +
      "matches this row's date and amount — likely the correct vendor. Confirm and fill in the GSTIN.";
    const suggestedPortalRow = portalRow('FTC-8842', {
      gstin: '32AAGFF3216M1ZL',
      narration: 'FABS TRADING COMPANY',
      invoiceDate: '2024-11-25',
      taxableValue: 5001.04,
    });
    const rows: GstResultRow[] = [
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('', { gstin: '', narration: 'Fabs Trading Company', invoiceDate: '2024-11-25', taxableValue: 5001.04 }),
        mismatchReason: 'blank_gstin_likely_matched',
        explanation,
        closestPortalRow: suggestedPortalRow,
      },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('', { gstin: '', narration: 'Still Blank Vendor' }),
        mismatchReason: 'blank_counterparty_gstin',
        explanation: 'GSTIN is blank in the register for this row — cannot be matched.',
      },
    ];

    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    const rowsOut = writeTableRows(actions);
    const titles = rowsOut.map((r) => r[0]);

    const blankGstinIdx = titles.indexOf('Blank GSTIN rows');
    const likelyMatchedIdx = rowsOut.findIndex(
      (r, i) =>
        r[0] === 'Blank GSTIN — likely matched (confirm & fill in GSTIN)' &&
        rowsOut[i + 1]?.[0] === 'Vendor Name (Books)',
    );
    expect(blankGstinIdx).toBeGreaterThan(-1);
    expect(likelyMatchedIdx).toBeGreaterThan(blankGstinIdx);

    const headerRow = rowsOut[likelyMatchedIdx + 1];
    expect(headerRow.slice(0, 4)).toEqual([
      'Vendor Name (Books)',
      'Invoice Date',
      'Taxable Value',
      'Suggested GSTIN',
    ]);
    const dataRow = rowsOut[likelyMatchedIdx + 2];
    expect(dataRow[0]).toBe('Fabs Trading Company');
    expect(dataRow[3]).toBe('32AAGFF3216M1ZL');
    expect(dataRow[4]).toBe('FABS TRADING COMPANY');
    expect(dataRow[5]).toBe('FTC-8842');

    // Never double-listed: the still-genuinely-blank row stays in "Blank GSTIN rows"; the
    // resolved one must not also appear there.
    const blankGstinSectionRows = rowsOut.slice(blankGstinIdx + 2, likelyMatchedIdx);
    expect(blankGstinSectionRows.some((r) => r[1] === 'Fabs Trading Company')).toBe(false);
    expect(blankGstinSectionRows.some((r) => r[1] === 'Still Blank Vendor')).toBe(true);
  });

  it('never silently drops a PR_ONLY row that lacks a recognized mismatch reason', () => {
    const rows: GstResultRow[] = [
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-NO-REASON'),
        // mismatchReason intentionally omitted — simulates a raw/older-shaped row.
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    const joined = JSON.stringify(writeTableRows(actions));
    expect(joined).toContain('INV-NO-REASON');
    expect(joined).toContain('Unclassified');
  });

  it('writes no sheet at all when every books row matched, even if the portal has its own extra rows (one-directional: nothing to report on the books side)', () => {
    const rows: GstResultRow[] = [
      {
        status: 'MATCHED',
        pass: 1,
        confidence: 1,
        itcAmount: 180,
        rcmFlag: false,
        registerRow: booksRow('INV-001'),
      },
      {
        // A portal row with no books counterpart — commonly another client's invoice
        // sharing the same portal export — must never by itself trigger a sheet, since
        // this reconciliation only ever reports on the books side.
        status: 'PORTAL_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        portalRow: portalRow('PORTAL-EXTRA'),
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });
    expect(actions).toHaveLength(0);
  });

  it('appends styled sheet-formatting actions after CREATE_SHEET + WRITE_TABLE (title/section navy bands, column widths, freeze panes)', () => {
    const rows: GstResultRow[] = [
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-BLANK-VALUE', { taxableValue: null }),
        mismatchReason: 'blank_taxable_value',
        explanation: 'No taxable value found in any rate column for this row.',
      },
    ];
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: rows,
    });

    expect(actions[0].type).toBe('CREATE_SHEET');
    expect(actions[1].type).toBe('WRITE_TABLE');
    const styling = actions.slice(2);
    expect(styling.length).toBeGreaterThan(0);
    expect(styling.every((a) => a.sheetName === 'Missed vs GSTR-2B')).toBe(true);

    const navyBands = styling.filter((a) => a.type === 'FORMAT_RANGE' && a.format?.fillColor === '#203764');
    // sheet title + (Blank taxable value: title+header) = 3 navy bands minimum.
    expect(navyBands.length).toBeGreaterThanOrEqual(3);
    expect(navyBands.every((a) => a.format?.bold === true && a.format?.fontColor === '#FFFFFF')).toBe(true);

    expect(styling.some((a) => a.type === 'SET_COLUMN_WIDTH')).toBe(true);
    expect(styling.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
    expect(styling.some((a) => a.type === 'CONDITIONAL_FORMAT')).toBe(true);
    // Taxable Value column gets a currency number format somewhere in the styling actions.
    expect(styling.some((a) => a.type === 'FORMAT_RANGE' && a.format?.numberFormat === '#,##0.00')).toBe(
      true,
    );

    // SUMMARY block's value column (B) is centered, alignment only — matches the real
    // sheet layout (15 metric rows: added "Possible RCM" and "Amended invoices" lines,
    // no "Portal-only"/"Unclassified" line here) -> Excel rows 5-19.
    const summaryCenter = styling.find((a) => a.type === 'FORMAT_RANGE' && a.range === 'B5:B19');
    expect(summaryCenter).toBeDefined();
    expect(summaryCenter?.format).toEqual({ horizontalAlignment: 'center' });
  });
});

describe('mapToSheetActions — full-report sheet inherits the same shared styling', () => {
  it('appends the same styling action types as the casual sheet, via the shared helper', () => {
    const actions = mapToSheetActions({
      sheetName: 'Recon PR VS GSTR2B',
      reconType: 'PR_VS_GSTR2B',
      summary: {
        total_pr_rows: 1,
        total_portal_rows: 1,
        total_ims_rows: 0,
        exact_matched: 1,
        partial_matched: 0,
        credit_notes: 0,
        pr_only: 0,
        portal_only: 0,
        ims_rejected: 0,
        ims_pending: 0,
        ims_auto_accept: 0,
        rcm_flagged: 0,
        itc_matched: 180,
        itc_at_risk: 0,
        rcm_payable: 0,
        itc_ims_rejected: 0,
        itc_ims_pending: 0,
        matched_exact: 1,
        matched_fallback: 0,
        mismatch_blank_gstin: 0,
        mismatch_blank_gstin_likely_matched: 0,
        mismatch_blank_taxable_value: 0,
        mismatch_ambiguous_rate_slab: 0,
        mismatch_amount: 0,
        mismatch_date: 0,
        mismatch_genuinely_missing: 0,
        gstin_mismatch_count: 0,
        amended_count: 0,
      },
      rows: [
        {
          pr_ref: 'PR:row_2',
          portal_ref: '2b:row_2',
          status: 'MATCHED',
          pass: 1,
          confidence: 1,
          difference: null,
          diff_type: null,
          itc_amount: 180,
          ims_status: null,
          rcm_flag: false,
          invoice_number: 'INV-001',
          gstin: '27AAAAA0000A1Z5',
          vendor_name: 'Some Vendor',
          mismatch_reason: null,
          explanation: null,
          books_sheet_name: 'PR',
          books_row: 2,
        },
      ],
      runAt: '2026-04-01T00:00:00.000Z',
    });

    expect(actions[0].type).toBe('CREATE_SHEET');
    expect(actions[1].type).toBe('WRITE_TABLE');
    const styling = actions.slice(2);

    const navyBands = styling.filter((a) => a.type === 'FORMAT_RANGE' && a.format?.fillColor === '#203764');
    // sheet title band + DETAIL section (title + header) = 3 minimum.
    expect(navyBands.length).toBeGreaterThanOrEqual(3);
    expect(styling.some((a) => a.type === 'SET_COLUMN_WIDTH')).toBe(true);
    expect(styling.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
    expect(styling.some((a) => a.type === 'CONDITIONAL_FORMAT')).toBe(true);
  });
});

/**
 * Shared fixture for the two alternate flat layouts: one row of every books-side
 * mismatch reason, one GSTIN-mismatch row, one matched row, and one portal-only row —
 * fed to BOTH mappers so a swapped direction (the bug this feature must never
 * reintroduce) would show up as the wrong function pulling the wrong rows.
 */
function buildDirectionalFixtureRows(): GstResultRow[] {
  return [
    {
      status: 'MATCHED',
      pass: 1,
      confidence: 1,
      itcAmount: 180,
      rcmFlag: false,
      registerRow: booksRow('INV-MATCHED'),
    },
    {
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('', { gstin: '', narration: 'Blank GSTIN Vendor' }),
      mismatchReason: 'blank_counterparty_gstin',
      explanation: 'GSTIN is blank.',
    },
    {
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('INV-AMBIGUOUS-SLAB', { narration: 'Ambiguous Slab Vendor' }),
      mismatchReason: 'ambiguous_rate_slab',
      explanation: 'Ambiguous rate slab.',
    },
    {
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('INV-AMT-MISMATCH', { narration: 'Amount Mismatch Vendor' }),
      mismatchReason: 'amount_mismatch',
      explanation: 'Amount differs.',
    },
    {
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('INV-DATE-MISMATCH', { narration: 'Date Mismatch Vendor' }),
      mismatchReason: 'date_mismatch',
      explanation: 'Date differs.',
    },
    {
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('INV-GENUINELY-MISSING', { narration: 'Genuinely Missing Vendor' }),
      mismatchReason: 'genuinely_missing',
      explanation: 'Not found anywhere in the portal.',
    },
    {
      status: 'GSTIN_MISMATCH',
      pass: null,
      confidence: 0.85,
      itcAmount: 0,
      rcmFlag: false,
      registerRow: booksRow('', { gstin: '32AIOPJ2231N1Z8', narration: 'Deva Steels' }),
      portalRow: portalRow('CBR/24-25/1691', { gstin: '33AIOPJ2231N1Z6', narration: 'DEVA STEELS' }),
      mismatchReason: 'gstin_mismatch_same_pan',
      explanation: 'Same vendor, different GSTIN registration.',
    },
    {
      status: 'PORTAL_ONLY',
      pass: null,
      confidence: 1,
      itcAmount: 0,
      rcmFlag: false,
      portalRow: portalRow('PORTAL-ONLY-1', { narration: 'Portal Only Vendor' }),
    },
  ];
}

describe('mapMissedBooksFlatToSheetActions — books-only flat layout', () => {
  it('pulls every books-side reason (PR_ONLY, any reason, plus GSTIN mismatch) and NEVER a PORTAL_ONLY or MATCHED row', () => {
    const actions = mapMissedBooksFlatToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      relativeTo: 'PR',
      resultRows: buildDirectionalFixtureRows(),
    });

    expect(actions[0]).toMatchObject({ type: 'CREATE_SHEET', sheetName: 'Missed vs GSTR-2B' });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    // Every books-side reason is present, in one flat table.
    expect(joined).toContain('Blank GSTIN Vendor');
    expect(joined).toContain('INV-AMBIGUOUS-SLAB');
    expect(joined).toContain('INV-AMT-MISMATCH');
    expect(joined).toContain('INV-DATE-MISMATCH');
    expect(joined).toContain('INV-GENUINELY-MISSING');
    expect(joined).toContain('Deva Steels'); // GSTIN-mismatch row, books-side vendor name
    expect(joined).toContain('32AIOPJ2231N1Z8'); // GSTIN-mismatch row's books-side GSTIN

    // Never the reverse direction (portal-only) or a matched row.
    expect(joined).not.toContain('PORTAL-ONLY-1');
    expect(joined).not.toContain('Portal Only Vendor');
    expect(joined).not.toContain('INV-MATCHED');

    // Single flat table — no reason-grouped section titles like the categorized layout.
    const sectionTitles = rowsOut.map((r) => r[0]).filter((c) => typeof c === 'string');
    expect(sectionTitles).not.toEqual(
      expect.arrayContaining(['Blank GSTIN rows', 'Ambiguous rate slab rows (needs CA review)']),
    );
    // Exactly one section header row (GSTIN | Vendor Name | ...).
    const headerRowCount = rowsOut.filter(
      (r) => r[0] === 'GSTIN' && r[1] === 'Vendor Name',
    ).length;
    expect(headerRowCount).toBe(1);
  });

  it('returns no actions when there are no books-side unmatched or GSTIN-mismatch rows', () => {
    const rows: GstResultRow[] = [
      { status: 'MATCHED', pass: 1, confidence: 1, itcAmount: 180, rcmFlag: false, registerRow: booksRow('INV-001') },
      { status: 'PORTAL_ONLY', pass: null, confidence: 1, itcAmount: 0, rcmFlag: false, portalRow: portalRow('PORTAL-1') },
    ];
    expect(
      mapMissedBooksFlatToSheetActions({
        sheetName: 'Missed vs GSTR-2B',
        portalLabel: 'GSTR-2B',
        booksLabel: 'Purchase Register',
        runAt: '2026-04-01T00:00:00.000Z',
        resultRows: rows,
      }),
    ).toEqual([]);
  });

  it('reuses the shared styling helper (navy bands, column widths, freeze panes, row banding)', () => {
    const actions = mapMissedBooksFlatToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
    });
    const styling = actions.slice(2);
    expect(styling.some((a) => a.type === 'FORMAT_RANGE' && a.format?.fillColor === '#203764')).toBe(true);
    expect(styling.some((a) => a.type === 'SET_COLUMN_WIDTH')).toBe(true);
    expect(styling.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
    expect(styling.some((a) => a.type === 'CONDITIONAL_FORMAT')).toBe(true);
  });
});

describe('mapPortalOnlyFlatToSheetActions — portal-only flat layout (reverse direction)', () => {
  it('pulls only PORTAL_ONLY rows and NEVER a books-side (PR_ONLY/GSTIN_MISMATCH) or MATCHED row', () => {
    const actions = mapPortalOnlyFlatToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      relativeTo: 'PR',
      resultRows: buildDirectionalFixtureRows(),
    });

    expect(actions[0]).toMatchObject({ type: 'CREATE_SHEET', sheetName: 'Missed vs GSTR-2B' });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    expect(joined).toContain('PORTAL-ONLY-1');
    expect(joined).toContain('Portal Only Vendor');

    // Never the reverse direction (any books-side reason) or the matched row.
    expect(joined).not.toContain('Blank GSTIN Vendor');
    expect(joined).not.toContain('INV-AMBIGUOUS-SLAB');
    expect(joined).not.toContain('INV-AMT-MISMATCH');
    expect(joined).not.toContain('INV-DATE-MISMATCH');
    expect(joined).not.toContain('INV-GENUINELY-MISSING');
    expect(joined).not.toContain('Deva Steels');
    expect(joined).not.toContain('CBR/24-25/1691');
    expect(joined).not.toContain('INV-MATCHED');

    const headerRowCount = rowsOut.filter(
      (r) => r[0] === 'GSTIN' && r[1] === 'Vendor Name',
    ).length;
    expect(headerRowCount).toBe(1);
  });

  it('returns no actions when there are no portal-only rows', () => {
    const rows: GstResultRow[] = [
      { status: 'MATCHED', pass: 1, confidence: 1, itcAmount: 180, rcmFlag: false, registerRow: booksRow('INV-001') },
      {
        status: 'PR_ONLY',
        pass: null,
        confidence: 1,
        itcAmount: 0,
        rcmFlag: false,
        registerRow: booksRow('INV-002'),
        mismatchReason: 'genuinely_missing',
      },
    ];
    expect(
      mapPortalOnlyFlatToSheetActions({
        sheetName: 'Missed vs GSTR-2B',
        portalLabel: 'GSTR-2B',
        booksLabel: 'Purchase Register',
        runAt: '2026-04-01T00:00:00.000Z',
        resultRows: rows,
      }),
    ).toEqual([]);
  });

  it('reuses the shared styling helper (navy bands, column widths, freeze panes, row banding)', () => {
    const actions = mapPortalOnlyFlatToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
    });
    const styling = actions.slice(2);
    expect(styling.some((a) => a.type === 'FORMAT_RANGE' && a.format?.fillColor === '#203764')).toBe(true);
    expect(styling.some((a) => a.type === 'SET_COLUMN_WIDTH')).toBe(true);
    expect(styling.some((a) => a.type === 'FREEZE_PANES')).toBe(true);
    expect(styling.some((a) => a.type === 'CONDITIONAL_FORMAT')).toBe(true);
  });
});

describe('isSales: true — column headers use sales terminology, not purchase wording', () => {
  it('mapMissedBooksToSheetActions: "Recipient GSTIN"/"Customer Name" columns, "same customer" wording, never "Vendor Name" or "same vendor"', () => {
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-1',
      portalLabel: 'GSTR-1',
      booksLabel: 'Sales Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
      isSales: true,
    });
    const rowsOut = writeTableRows(actions);
    const joined = JSON.stringify(rowsOut);

    expect(joined).toContain('Recipient GSTIN');
    expect(joined).toContain('Customer Name');
    expect(joined).toContain('same customer, different registration');
    expect(joined).not.toContain('Vendor Name');
    expect(joined).not.toContain('same vendor, different registration');

    // Blank-GSTIN section title/summary line use "recipient GSTIN" too.
    expect(joined).toContain('Blank recipient GSTIN rows');
    expect(joined).toContain('Blank recipient GSTIN in books');
  });

  it('mapMissedBooksFlatToSheetActions: sales headers on the books-only flat layout', () => {
    const actions = mapMissedBooksFlatToSheetActions({
      sheetName: 'Missed vs GSTR-1 — Books Only',
      portalLabel: 'GSTR-1',
      booksLabel: 'Sales Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
      isSales: true,
    });
    const joined = JSON.stringify(writeTableRows(actions));
    expect(joined).toContain('Recipient GSTIN');
    expect(joined).toContain('Customer Name');
    expect(joined).not.toContain('Vendor Name');
  });

  it('mapPortalOnlyFlatToSheetActions: sales headers on the portal-only flat layout', () => {
    const actions = mapPortalOnlyFlatToSheetActions({
      sheetName: 'Missed vs GSTR-1 — Portal Only',
      portalLabel: 'GSTR-1',
      booksLabel: 'Sales Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
      isSales: true,
    });
    const joined = JSON.stringify(writeTableRows(actions));
    expect(joined).toContain('Recipient GSTIN');
    expect(joined).toContain('Customer Name');
    expect(joined).not.toContain('Vendor Name');
  });

  it('mapToSheetActions (full report): sales headers on the DETAIL section', () => {
    const actions = mapToSheetActions({
      sheetName: 'Recon SALES VS GSTR1',
      reconType: 'SALES_VS_GSTR1',
      isSales: true,
      summary: {
        total_pr_rows: 1,
        total_portal_rows: 1,
        total_ims_rows: 0,
        exact_matched: 1,
        partial_matched: 0,
        credit_notes: 0,
        pr_only: 0,
        portal_only: 0,
        ims_rejected: 0,
        ims_pending: 0,
        ims_auto_accept: 0,
        rcm_flagged: 0,
        itc_matched: 180,
        itc_at_risk: 0,
        rcm_payable: 0,
        itc_ims_rejected: 0,
        itc_ims_pending: 0,
        matched_exact: 1,
        matched_fallback: 0,
        mismatch_blank_gstin: 0,
        mismatch_blank_gstin_likely_matched: 0,
        mismatch_blank_taxable_value: 0,
        mismatch_ambiguous_rate_slab: 0,
        mismatch_amount: 0,
        mismatch_date: 0,
        mismatch_genuinely_missing: 0,
        gstin_mismatch_count: 0,
        amended_count: 0,
      },
      rows: [
        {
          pr_ref: 'SR:row_2',
          portal_ref: 'GSTR1:row_2',
          status: 'MATCHED',
          pass: 1,
          confidence: 1,
          difference: null,
          diff_type: null,
          itc_amount: 180,
          ims_status: null,
          rcm_flag: false,
          invoice_number: 'INV-001',
          gstin: '27AAAAA0000A1Z5',
          vendor_name: 'Some Customer',
          mismatch_reason: null,
          explanation: null,
          books_sheet_name: 'SR',
          books_row: 2,
        },
      ],
      runAt: '2026-04-01T00:00:00.000Z',
    });

    const rowsOut = writeTableRows(actions);
    const detailHeaderRow = rowsOut.find((r) => r[0] === 'Status');
    expect(detailHeaderRow?.slice(2, 4)).toEqual(['Recipient GSTIN', 'Customer Name']);
  });

  it('purchase (isSales omitted) keeps the original wording — never leaks sales terminology', () => {
    const actions = mapMissedBooksToSheetActions({
      sheetName: 'Missed vs GSTR-2B',
      portalLabel: 'GSTR-2B',
      booksLabel: 'Purchase Register',
      runAt: '2026-04-01T00:00:00.000Z',
      resultRows: buildDirectionalFixtureRows(),
    });
    const joined = JSON.stringify(writeTableRows(actions));
    expect(joined).toContain('Vendor Name');
    expect(joined).toContain('same vendor, different registration');
    expect(joined).not.toContain('Recipient GSTIN');
    expect(joined).not.toContain('Customer Name');
    expect(joined).not.toContain('same customer, different registration');
  });
});
