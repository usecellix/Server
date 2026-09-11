import { gstMatch } from './gst-match.tool';
import { normalizeInvoiceNumber } from './normalize-invoice';
import { NormalizedInvoiceRow } from '../types/domain-tool.types';

let rowCounter = 0;

function row(overrides: Partial<NormalizedInvoiceRow> = {}): NormalizedInvoiceRow {
  rowCounter += 1;
  const invoiceNumber = overrides.invoiceNumber ?? `INV-${rowCounter}`;
  return {
    gstin: '32AAAAA0000A1Z5',
    invoiceNumber,
    normalizedInvoiceNumber: normalizeInvoiceNumber(invoiceNumber),
    invoiceDate: '2026-04-15',
    taxableValue: 10000,
    taxAmount: 1800,
    igst: 0,
    cgst: 900,
    sgst: 900,
    narration: '',
    documentType: 'invoice',
    sourceRowRef: {
      documentType: 'workbook',
      documentId: 'synth-register',
      rowOrLine: rowCounter,
    },
    ...overrides,
  };
}

describe('Pass 1b — fallback match (no Invoice Number column)', () => {
  it('recovers a GSTIN + date + amount match when the books sheet has no invoice numbers', () => {
    const books = row({ invoiceNumber: '', normalizedInvoiceNumber: '' });
    const portal = row({
      invoiceNumber: 'PORTAL-REF-1',
      normalizedInvoiceNumber: normalizeInvoiceNumber('PORTAL-REF-1'),
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 99 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
      booksHasInvoiceNumberColumn: false,
    });

    expect(result.data.resultRows).toHaveLength(1);
    expect(result.data.resultRows[0].status).toBe('MATCHED');
    expect(result.data.resultRows[0].pass).toBe(2);
    expect(result.data.resultRows[0].diffType).toBe('FALLBACK_NO_INVOICE_NUMBER');
  });

  it('does NOT run the fallback when booksHasInvoiceNumberColumn is not explicitly false', () => {
    const books = row({ invoiceNumber: '', normalizedInvoiceNumber: '' });
    const portal = row({
      invoiceNumber: 'PORTAL-REF-1',
      normalizedInvoiceNumber: normalizeInvoiceNumber('PORTAL-REF-1'),
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 99 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
    });

    expect(result.data.resultRows[0].status).not.toBe('MATCHED');
  });

  it('respects amount tolerance (±₹1 / ±0.5% by default) and exact date', () => {
    const books = row({ invoiceNumber: '', normalizedInvoiceNumber: '', taxableValue: 10000 });
    const portal = row({
      invoiceNumber: 'PORTAL-REF-2',
      normalizedInvoiceNumber: normalizeInvoiceNumber('PORTAL-REF-2'),
      taxableValue: 10000.4, // within ±₹1
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 100 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
      booksHasInvoiceNumberColumn: false,
    });

    expect(result.data.resultRows[0].status).toBe('MATCHED');
  });
});

describe('finalizeUnmatched — mismatch reason diagnostics', () => {
  it('tags a blank-GSTIN row as blank_counterparty_gstin (RCM detection off, as in casual/missed_books_only mode)', () => {
    const books = row({ gstin: '' });
    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [],
      settings: { detectRcm: false },
    });

    const pr = result.data.resultRows[0];
    expect(pr.status).toBe('PR_ONLY');
    expect(pr.mismatchReason).toBe('blank_counterparty_gstin');
    expect(pr.explanation).toMatch(/GSTIN is blank/i);
  });

  it('tags a blank-taxable-value row as blank_taxable_value', () => {
    const books = row({ taxableValue: null });
    const result = gstMatch({ purchaseRegister: [books], gstr2b: [] });

    const pr = result.data.resultRows[0];
    expect(pr.status).toBe('PR_ONLY');
    expect(pr.mismatchReason).toBe('blank_taxable_value');
    expect(pr.explanation).toMatch(/no taxable value/i);
  });

  it('tags a row flagged ambiguousRateSlab as ambiguous_rate_slab, taking priority over other checks', () => {
    const books = row({
      taxableValue: null,
      ambiguousRateSlab: true,
      ambiguousRateSlabDetail: 'Multiple rate-slab columns are populated (...) but none matches.',
    });
    const result = gstMatch({ purchaseRegister: [books], gstr2b: [] });

    const pr = result.data.resultRows[0];
    expect(pr.status).toBe('PR_ONLY');
    expect(pr.mismatchReason).toBe('ambiguous_rate_slab');
    expect(pr.explanation).toMatch(/multiple rate-slab columns/i);
  });

  it('tags a row whose GSTIN never appears in the portal as gstin_not_in_portal', () => {
    const books = row({ gstin: '32AAAAA0000A1Z5' });
    const portal = row({ gstin: '27ZZZZZ0000Z1Z5', invoiceNumber: 'OTHER-1' });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });

    const pr = result.data.resultRows.find((r) => r.status === 'PR_ONLY');
    expect(pr?.mismatchReason).toBe('gstin_not_in_portal');
    expect(pr?.explanation).toContain('32AAAAA0000A1Z5');
  });

  it('tags same-GSTIN-and-date-but-different-amount as amount_mismatch with fieldDiff', () => {
    const books = row({ invoiceNumber: 'BOOKS-VERY-DIFFERENT-NUMBER-A', taxableValue: 10000, invoiceDate: '2026-04-15' });
    const portal = row({
      invoiceNumber: 'PORTAL-COMPLETELY-UNLIKE-NUMBER-B',
      taxableValue: 15000,
      invoiceDate: '2026-04-15',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 200 },
    });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });

    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.status).toBe('PR_ONLY');
    expect(pr?.mismatchReason).toBe('amount_mismatch');
    expect(pr?.fieldDiff).toEqual([
      { field: 'taxableValue', booksValue: 10000, portalValue: 15000 },
    ]);
    expect(pr?.closestPortalRow?.taxableValue).toBe(15000);
    expect(pr?.explanation).toMatch(/amount differs/i);
  });

  it('tags same-GSTIN-and-amount-but-different-date as date_mismatch with fieldDiff', () => {
    const books = row({ invoiceNumber: 'BOOKS-VERY-DIFFERENT-NUMBER-C', taxableValue: 10000, invoiceDate: '2026-04-15' });
    const portal = row({
      invoiceNumber: 'PORTAL-COMPLETELY-UNLIKE-NUMBER-D',
      taxableValue: 10000,
      invoiceDate: '2026-05-20',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 201 },
    });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });

    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.status).toBe('PR_ONLY');
    expect(pr?.mismatchReason).toBe('date_mismatch');
    expect(pr?.fieldDiff).toEqual([
      { field: 'invoiceDate', booksValue: '2026-04-15', portalValue: '2026-05-20' },
    ]);
    expect(pr?.explanation).toMatch(/date differs/i);
  });

  it('tags a row with no plausible candidate anywhere as genuinely_missing', () => {
    const books = row({ invoiceNumber: 'BOOKS-X', taxableValue: 10000, invoiceDate: '2026-04-15' });
    const portal = row({
      invoiceNumber: 'PORTAL-Y',
      taxableValue: 99999,
      invoiceDate: '2026-01-01',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 202 },
    });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });

    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.mismatchReason).toBe('genuinely_missing');
  });

  it('keeps MATCHED/PARTIAL/CREDIT_NOTE rows untouched by mismatch diagnosis', () => {
    const books = row({ invoiceNumber: 'SAME-INVOICE-NUMBER' });
    const portal = row({
      invoiceNumber: 'SAME-INVOICE-NUMBER',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 300 },
    });
    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });

    expect(result.data.resultRows[0].status).toBe('MATCHED');
    expect(result.data.resultRows[0].mismatchReason).toBeUndefined();
  });
});

/**
 * Regression coverage for the PAN-based GSTIN-mismatch pass, using the exact 9 real
 * pairs found auditing "PR vs Missed PR.xlsx" for May 2024: Deva Steels (GSTIN
 * 32AIOPJ2231N1Z8 in books vs 33AIOPJ2231N1Z6 in GSTR-2B, ×6), Dalmia Cement (
 * 32AADCA9414C1Z8 vs 33AADCA9414C1Z6, ×2), JSW Cement (37AABCJ6731B1ZV vs
 * 29AABCJ6731B1ZS, ×1). All three pairs share the same PAN (characters 3-12) despite
 * different GSTINs. Two of the Deva Steels pairs additionally have a genuine date
 * discrepancy (books date differs from the portal invoice date by 8-14 days) — real
 * data, not a parsing artifact — so the pass must match on PAN + amount even when the
 * date differs, while still preferring an exact date match when one exists.
 */
describe('runPanCrossGstinMatch — same PAN, different GSTIN registration', () => {
  const realPairs = [
    // [label, booksGstin, portalGstin, date, portalDate, amount]
    { label: 'Dalmia #1', booksGstin: '32AADCA9414C1Z8', portalGstin: '33AADCA9414C1Z6', date: '2024-05-16', portalDate: '2024-05-16', amount: 75000, portalInvoice: '2401043632' },
    { label: 'Dalmia #2', booksGstin: '32AADCA9414C1Z8', portalGstin: '33AADCA9414C1Z6', date: '2024-05-16', portalDate: '2024-05-16', amount: 18750, portalInvoice: '2401043637' },
    { label: 'Deva #1', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-09', portalDate: '2024-05-09', amount: 120780.4, portalInvoice: 'CBR/24-25/1691' },
    { label: 'Deva #2', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-24', portalDate: '2024-05-24', amount: 35655.87, portalInvoice: 'CBR/24-25/2436' },
    { label: 'Deva #3', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-30', portalDate: '2024-05-30', amount: 22921.63, portalInvoice: 'CBR/24-25/2731' },
    { label: 'Deva #4', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-30', portalDate: '2024-05-30', amount: 13073.82, portalInvoice: 'CBR/24-25/2732' },
    // Genuine books-vs-portal date discrepancy on top of the GSTIN mismatch:
    { label: 'Deva #5 (date-shifted)', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-16', portalDate: '2024-05-24', amount: 36198.41, portalInvoice: 'CBR/24-25/2437' },
    { label: 'Deva #6 (date-shifted)', booksGstin: '32AIOPJ2231N1Z8', portalGstin: '33AIOPJ2231N1Z6', date: '2024-05-16', portalDate: '2024-05-30', amount: 80650.18, portalInvoice: 'CBR/24-25/2730' },
    { label: 'JSW Cement', booksGstin: '37AABCJ6731B1ZV', portalGstin: '29AABCJ6731B1ZS', date: '2024-05-14', portalDate: '2024-05-14', amount: 88541.25, portalInvoice: 'KA2402025552' },
  ];

  it('reclassifies all 9 real pairs out of genuinely_missing/gstin_not_in_portal/portal_only into GSTIN_MISMATCH', () => {
    const books = realPairs.map((p) =>
      row({ gstin: p.booksGstin, invoiceNumber: '', normalizedInvoiceNumber: '', invoiceDate: p.date, taxableValue: p.amount }),
    );
    const portal = realPairs.map((p) =>
      row({
        gstin: p.portalGstin,
        invoiceNumber: p.portalInvoice,
        normalizedInvoiceNumber: normalizeInvoiceNumber(p.portalInvoice),
        invoiceDate: p.portalDate,
        taxableValue: p.amount,
        sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 1000 },
      }),
    );

    const result = gstMatch({
      purchaseRegister: books,
      gstr2b: portal,
      settings: { detectRcm: false },
      booksHasInvoiceNumberColumn: false,
    });

    const mismatchRows = result.data.resultRows.filter((r) => r.status === 'GSTIN_MISMATCH');
    expect(mismatchRows).toHaveLength(9);
    expect(result.data.resultRows.filter((r) => r.status === 'PR_ONLY')).toHaveLength(0);
    expect(result.data.resultRows.filter((r) => r.status === 'PORTAL_ONLY')).toHaveLength(0);

    for (const p of realPairs) {
      const match = mismatchRows.find(
        (r) => r.registerRow?.gstin === p.booksGstin && r.registerRow?.taxableValue === p.amount,
      );
      expect(match).toBeDefined();
      expect(match?.portalRow?.gstin).toBe(p.portalGstin);
      expect(match?.mismatchReason).toBe('gstin_mismatch_same_pan');
      expect(match?.explanation).toContain(p.booksGstin);
      expect(match?.explanation).toContain(p.portalGstin);
    }
  });

  it('explanation names both GSTINs and the PAN, matching the exact real example', () => {
    const books = row({
      gstin: '32AIOPJ2231N1Z8',
      invoiceNumber: '',
      normalizedInvoiceNumber: '',
      invoiceDate: '2024-05-09',
      taxableValue: 120780.4,
    });
    const portal = row({
      gstin: '33AIOPJ2231N1Z6',
      invoiceNumber: 'CBR/24-25/1691',
      normalizedInvoiceNumber: normalizeInvoiceNumber('CBR/24-25/1691'),
      invoiceDate: '2024-05-09',
      taxableValue: 120780.4,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 1001 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
      settings: { detectRcm: false },
      booksHasInvoiceNumberColumn: false,
    });

    const mismatch = result.data.resultRows[0];
    expect(mismatch.status).toBe('GSTIN_MISMATCH');
    expect(mismatch.explanation).toBe(
      'Same vendor (PAN AIOPJ2231N), same date and amount, but booked under GSTIN 32AIOPJ2231N1Z8 in your register ' +
        'vs GSTIN 33AIOPJ2231N1Z6 in GSTR-2B — check which registration this vendor actually used for this invoice.',
    );
  });

  it('when dates differ, the explanation says so explicitly rather than claiming "same date"', () => {
    const books = row({
      gstin: '32AIOPJ2231N1Z8',
      invoiceNumber: '',
      normalizedInvoiceNumber: '',
      invoiceDate: '2024-05-16',
      taxableValue: 36198.41,
    });
    const portal = row({
      gstin: '33AIOPJ2231N1Z6',
      invoiceNumber: 'CBR/24-25/2437',
      normalizedInvoiceNumber: normalizeInvoiceNumber('CBR/24-25/2437'),
      invoiceDate: '2024-05-24',
      taxableValue: 36198.41,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 1002 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
      settings: { detectRcm: false },
      booksHasInvoiceNumberColumn: false,
    });

    const mismatch = result.data.resultRows[0];
    expect(mismatch.status).toBe('GSTIN_MISMATCH');
    expect(mismatch.explanation).toContain('the invoice date also differs');
    expect(mismatch.explanation).toContain('books 2024-05-16');
    expect(mismatch.explanation).toContain('portal 2024-05-24');
  });

  it('does not fire when the GSTIN is merely blank (that stays blank_counterparty_gstin)', () => {
    const books = row({ gstin: '', taxableValue: 5000 });
    const result = gstMatch({ purchaseRegister: [books], gstr2b: [], settings: { detectRcm: false } });
    expect(result.data.resultRows[0].status).toBe('PR_ONLY');
    expect(result.data.resultRows[0].mismatchReason).toBe('blank_counterparty_gstin');
  });

  it('does not fire when amount does not match even if PAN matches', () => {
    const books = row({ gstin: '32AIOPJ2231N1Z8', invoiceNumber: '', normalizedInvoiceNumber: '', taxableValue: 1000 });
    const portal = row({
      gstin: '33AIOPJ2231N1Z6',
      invoiceNumber: 'UNRELATED',
      taxableValue: 99999,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 1003 },
    });
    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portal],
      settings: { detectRcm: false },
      booksHasInvoiceNumberColumn: false,
    });
    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.status).toBe('PR_ONLY');
    expect(pr?.status).not.toBe('GSTIN_MISMATCH');
  });

  it('never interferes with normal exact GSTIN+invoice matching', () => {
    const books = row({ gstin: '32AIOPJ2231N1Z8', invoiceNumber: 'SAME-INV' });
    const portal = row({
      gstin: '32AIOPJ2231N1Z8',
      invoiceNumber: 'SAME-INV',
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 1004 },
    });
    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal] });
    expect(result.data.resultRows[0].status).toBe('MATCHED');
  });
});

describe('row-accounting invariant', () => {
  it('every books row ends up matched or diagnosed — matched + diagnosed === total books rows', () => {
    const books = [
      row({}), // exact match
      row({ gstin: '' }), // blank gstin
      row({ taxableValue: null }), // blank value
      row({ gstin: '99BBBBB0000B1Z5', invoiceNumber: 'NO-PORTAL-MATCH' }), // gstin not in portal
    ];
    const portal = [
      row({
        invoiceNumber: books[0].invoiceNumber,
        sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 500 },
      }),
    ];

    const result = gstMatch({ purchaseRegister: books, gstr2b: portal, mode: 'purchase' });
    const accountedFor = result.data.resultRows.filter((r) => r.registerRow).length;
    expect(accountedFor).toBe(books.length);
  });

  it('holds for sales mode too', () => {
    const books = [row({ supplyCategory: 'b2c', gstin: '' }), row({})];
    const portal = [
      row({
        invoiceNumber: books[1].invoiceNumber,
        sourceRowRef: { documentType: 'gstr1', documentId: 'gstr1', rowOrLine: 600 },
      }),
    ];

    const result = gstMatch({ purchaseRegister: books, gstr2b: portal, mode: 'sales' });
    const accountedFor = result.data.resultRows.filter((r) => r.registerRow).length;
    expect(accountedFor).toBe(books.length);
  });
});

/**
 * Regression coverage for the blank-GSTIN double-counting bug: a books row with no GSTIN
 * never gets compared against portal data (GSTIN is the primary key), but its real
 * counterpart invoice still sits in the portal file and shows up separately as
 * portal_only — so the same real-world discrepancy is counted and shown twice. Confirmed
 * on the real file: Fabs Trading Company, ₹5,001.04 on 25-Nov-24, appeared both as a
 * blank-GSTIN books row and a portal-only row for the same invoice.
 */
describe('resolveLikelyBlankGstinMatches — blank-GSTIN rows resolved against portal_only', () => {
  it('reclassifies a blank-GSTIN row uniquely matched on date + amount, and removes its portal_only counterpart', () => {
    const books = row({
      gstin: '',
      narration: 'Fabs Trading Company',
      invoiceDate: '2024-11-25',
      taxableValue: 5001.04,
    });
    const portal = row({
      gstin: '32AAGFF3216M1ZL',
      narration: 'FABS TRADING COMPANY',
      invoiceNumber: 'FTC-8842',
      invoiceDate: '2024-11-25',
      taxableValue: 5001.04,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 900 },
    });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal], settings: { detectRcm: false } });

    expect(result.data.resultRows).toHaveLength(1);
    const only = result.data.resultRows[0];
    expect(only.status).toBe('PR_ONLY');
    expect(only.mismatchReason).toBe('blank_gstin_likely_matched');
    expect(only.closestPortalRow?.gstin).toBe('32AAGFF3216M1ZL');
    expect(only.closestPortalRow?.invoiceNumber).toBe('FTC-8842');
    expect(only.explanation).toBe(
      'GSTIN is blank in the register, but a portal invoice from FABS TRADING COMPANY (GSTIN 32AAGFF3216M1ZL) ' +
        "matches this row's date and amount — likely the correct vendor. Confirm and fill in the GSTIN.",
    );
    // The portal row must not also appear as a separate PORTAL_ONLY result — same invoice, not two.
    expect(result.data.resultRows.some((r) => r.status === 'PORTAL_ONLY')).toBe(false);
  });

  it('leaves both sides unchanged when zero candidates match', () => {
    const books = row({ gstin: '', invoiceDate: '2024-11-25', taxableValue: 5001.04 });
    const portal = row({
      gstin: '32AAGFF3216M1ZL',
      invoiceDate: '2024-11-25',
      taxableValue: 9999.99, // different amount — no candidate
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 901 },
    });

    const result = gstMatch({ purchaseRegister: [books], gstr2b: [portal], settings: { detectRcm: false } });
    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.mismatchReason).toBe('blank_counterparty_gstin');
    expect(result.data.resultRows.some((r) => r.status === 'PORTAL_ONLY')).toBe(true);
  });

  it('leaves both sides unchanged when multiple candidates match ambiguously', () => {
    const books = row({ gstin: '', invoiceDate: '2024-11-25', taxableValue: 5001.04 });
    const portalA = row({
      gstin: '32AAGFF3216M1ZL',
      invoiceDate: '2024-11-25',
      taxableValue: 5001.04,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 902 },
    });
    const portalB = row({
      gstin: '32BBGFF3216M1ZM',
      invoiceDate: '2024-11-25',
      taxableValue: 5001.04,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 903 },
    });

    const result = gstMatch({
      purchaseRegister: [books],
      gstr2b: [portalA, portalB],
      settings: { detectRcm: false },
    });
    const pr = result.data.resultRows.find((r) => r.registerRow);
    expect(pr?.mismatchReason).toBe('blank_counterparty_gstin');
    expect(result.data.resultRows.filter((r) => r.status === 'PORTAL_ONLY')).toHaveLength(2);
  });

  it('the row-accounting invariant holds: blank_gstin_likely_matched + remaining blank_counterparty_gstin === original blank-GSTIN count, and portal_only drops by the number resolved', () => {
    const resolvable = [
      row({ gstin: '', invoiceDate: '2024-11-25', taxableValue: 1000, narration: 'Vendor A' }),
      row({ gstin: '', invoiceDate: '2024-11-26', taxableValue: 2000, narration: 'Vendor B' }),
    ];
    const unresolvable = row({ gstin: '', invoiceDate: '2024-11-27', taxableValue: 3000, narration: 'Vendor C' });
    const books = [...resolvable, unresolvable];

    const portalMatches = resolvable.map((b, i) =>
      row({
        gstin: `32AAAA${i}000A1Z5`,
        invoiceDate: b.invoiceDate,
        taxableValue: b.taxableValue as number,
        sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 910 + i },
      }),
    );
    const unrelatedPortalOnly = row({
      gstin: '32ZZZZZ0000Z1Z5',
      invoiceDate: '2024-12-01',
      taxableValue: 9999,
      sourceRowRef: { documentType: 'gstr2b', documentId: '2b', rowOrLine: 920 },
    });
    const portal = [...portalMatches, unrelatedPortalOnly];

    const originalBlankGstinCount = books.length; // all 3 books rows are blank-GSTIN
    const originalPortalOnlyCount = portal.length; // none of these match anything else

    const result = gstMatch({ purchaseRegister: books, gstr2b: portal, settings: { detectRcm: false } });

    const likelyMatched = result.data.resultRows.filter((r) => r.mismatchReason === 'blank_gstin_likely_matched');
    const stillBlank = result.data.resultRows.filter((r) => r.mismatchReason === 'blank_counterparty_gstin');
    const portalOnly = result.data.resultRows.filter((r) => r.status === 'PORTAL_ONLY');

    expect(likelyMatched.length + stillBlank.length).toBe(originalBlankGstinCount);
    expect(likelyMatched).toHaveLength(2);
    expect(stillBlank).toHaveLength(1);
    expect(portalOnly).toHaveLength(originalPortalOnlyCount - 2);
  });
});
