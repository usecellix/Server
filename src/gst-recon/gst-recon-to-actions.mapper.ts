import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';
import {
  buildReconSheetStylingActions,
  StyledSheetBuilder,
} from '../excel-ai/utils/styled-recon-sheet.util';
import { MismatchReason, SourceRef } from '../domain-tools/types/domain-tool.types';
import { GstResultRow } from '../domain-tools/gst/match-passes';
import { ItcComputeOutput } from '../domain-tools/gst/itc-compute.tool';
import { Gstr3bVs2bOutput } from '../domain-tools/gst/gstr3b-vs-2b.tool';

export interface ReconSummaryBlock {
  total_pr_rows: number;
  total_portal_rows: number;
  total_ims_rows: number;
  exact_matched: number;
  partial_matched: number;
  credit_notes: number;
  pr_only: number;
  portal_only: number;
  ims_rejected: number;
  ims_pending: number;
  ims_auto_accept: number;
  rcm_flagged: number;
  itc_matched: number;
  itc_at_risk: number;
  rcm_payable: number;
  itc_ims_rejected: number;
  itc_ims_pending: number;
  cross_gstin_exception_count?: number;
  /** Matched breakdown by pass — never collapse this into a single "matched" number. */
  matched_exact: number;
  matched_fallback: number;
  /** Unmatched-in-books breakdown by reason (mirrors MismatchReason) — never a flat pr_only count alone. */
  mismatch_blank_gstin: number;
  /** Blank-GSTIN rows resolved against a unique portal_only candidate on date+amount — removed from portal_only, not double-counted. */
  mismatch_blank_gstin_likely_matched: number;
  mismatch_blank_taxable_value: number;
  mismatch_ambiguous_rate_slab: number;
  mismatch_amount: number;
  mismatch_date: number;
  mismatch_genuinely_missing: number;
  /** Same vendor (PAN), same invoice, booked under a different GSTIN registration on each side — its own bucket, not pr_only or portal_only. */
  gstin_mismatch_count: number;
}

export interface ReconRowDto {
  pr_ref: string | null;
  portal_ref: string | null;
  status: string;
  pass: number | null;
  confidence: number;
  difference: string | null;
  diff_type: string | null;
  itc_amount: number;
  ims_status: string | null;
  rcm_flag: boolean;
  invoice_number: string | null;
  gstin: string | null;
  vendor_name: string | null;
  mismatch_reason: MismatchReason | null;
  explanation: string | null;
}

function refLabel(ref?: SourceRef): string | null {
  if (!ref) return null;
  return `${ref.documentId}:row_${ref.rowOrLine}`;
}

/** Reasons that both mean "no plausible candidate anywhere" — grouped as one "genuinely missing" bucket for reporting. */
const GENUINELY_MISSING_REASONS: MismatchReason[] = ['gstin_not_in_portal', 'genuinely_missing'];

export function buildSummary(
  resultRows: GstResultRow[],
  itc: ItcComputeOutput,
  prCount: number,
  portalCount: number,
  imsCount: number,
): ReconSummaryBlock {
  const count = (status: string) => resultRows.filter((r) => r.status === status).length;
  const countReason = (...reasons: MismatchReason[]) =>
    resultRows.filter((r) => r.mismatchReason && reasons.includes(r.mismatchReason)).length;
  return {
    total_pr_rows: prCount,
    total_portal_rows: portalCount,
    total_ims_rows: imsCount,
    exact_matched: count('MATCHED'),
    partial_matched: count('PARTIAL') + count('IMS_AUTO_ACCEPT'),
    credit_notes: count('CREDIT_NOTE'),
    pr_only: count('PR_ONLY'),
    portal_only: count('PORTAL_ONLY') + count('IMS_ONLY'),
    ims_rejected: count('IMS_REJECTED'),
    ims_pending: count('IMS_PENDING'),
    ims_auto_accept: resultRows.filter(
      (r) => r.imsStatus === 'AutoAccepted' || r.status === 'IMS_AUTO_ACCEPT',
    ).length,
    rcm_flagged: count('RCM'),
    itc_matched: itc.totalItcClaimable,
    itc_at_risk: itc.totalItcAtRisk,
    rcm_payable: itc.totalRcmPayable,
    itc_ims_rejected: itc.totalImsRejected,
    itc_ims_pending: itc.totalImsPending,
    matched_exact: resultRows.filter((r) => r.status === 'MATCHED' && r.pass === 1).length,
    matched_fallback: resultRows.filter((r) => r.status === 'MATCHED' && r.pass === 2).length,
    mismatch_blank_gstin: countReason('blank_counterparty_gstin'),
    mismatch_blank_gstin_likely_matched: countReason('blank_gstin_likely_matched'),
    mismatch_blank_taxable_value: countReason('blank_taxable_value'),
    mismatch_ambiguous_rate_slab: countReason('ambiguous_rate_slab'),
    mismatch_amount: countReason('amount_mismatch'),
    mismatch_date: countReason('date_mismatch'),
    mismatch_genuinely_missing: countReason(...GENUINELY_MISSING_REASONS),
    gstin_mismatch_count: count('GSTIN_MISMATCH'),
  };
}

export function mapResultRows(resultRows: GstResultRow[]): ReconRowDto[] {
  return resultRows.map((r) => {
    const row = r.registerRow ?? r.portalRow ?? r.imsRow;
    return {
      pr_ref: refLabel(r.registerRow?.sourceRowRef),
      portal_ref: refLabel(r.portalRow?.sourceRowRef ?? r.imsRow?.sourceRowRef),
      status: r.status,
      pass: r.pass,
      confidence: r.confidence,
      difference: r.difference ?? null,
      diff_type: r.diffType ?? null,
      itc_amount: r.itcAmount,
      ims_status: r.imsStatus ?? null,
      rcm_flag: r.rcmFlag,
      invoice_number: row?.invoiceNumber ?? null,
      gstin: row?.gstin ?? null,
      vendor_name: row?.narration || null,
      mismatch_reason: r.mismatchReason ?? null,
      explanation: r.explanation ?? null,
    };
  });
}

export function mapToSheetActions(params: {
  sheetName: string;
  period?: string;
  gstin?: string;
  clientName?: string;
  financialYear?: string;
  operatorName?: string;
  firmName?: string;
  reconType: string;
  summary: ReconSummaryBlock;
  rows: ReconRowDto[];
  runAt: string;
  relativeTo?: string;
  isSales?: boolean;
}): SheetActionPayload[] {
  const {
    sheetName,
    period,
    gstin,
    clientName,
    financialYear,
    operatorName,
    firmName,
    reconType,
    summary,
    rows,
    runAt,
    relativeTo,
    isSales,
  } = params;

  const taxMatchedLabel = isSales ? 'Tax liability matched' : 'ITC matched';
  const taxAtRiskLabel = isSales ? 'Tax liability at risk' : 'ITC at risk';
  const booksLabel = isSales ? 'Sales Register Rows' : 'Purchase Register Rows';
  const booksOnlyLabel = isSales ? 'Books Only' : 'PR Only';
  const reconTypeLabel = isSales
    ? 'Sales Register vs GSTR-1'
    : reconType.replace(/_/g, ' ');

  const headers = [
    'Status',
    'Pass',
    'GSTIN',
    'Vendor Name',
    'Invoice No',
    isSales ? 'Tax Amount' : 'ITC Amount',
    'Confidence',
    'Difference',
    'Diff Type',
    'IMS Status',
    'RCM',
    'Books Ref',
    'Portal Ref',
  ];

  const tableRows: unknown[][] = rows.map((r) => [
    r.status,
    r.pass ?? '',
    r.gstin ?? '',
    r.vendor_name ?? '',
    r.invoice_number ?? '',
    r.itc_amount,
    r.confidence,
    r.difference ?? '',
    r.diff_type ?? '',
    r.ims_status ?? '',
    r.rcm_flag ? 'Y' : '',
    r.pr_ref ?? '',
    r.portal_ref ?? '',
  ]);

  const width = headers.length;
  const builder = new StyledSheetBuilder(width);
  const sheetTitleRow = builder.pushRow(['CELLIX GST Reconciliation Report']);
  builder.pushRow(['Client Name:', clientName ?? '']);
  builder.pushRow(['Client GSTIN:', gstin ?? '']);
  builder.pushRow(['Reconciliation Type:', reconTypeLabel]);
  builder.pushRow(['Period:', `${period ?? ''}${financialYear ? ` (FY ${financialYear})` : ''}`]);
  builder.pushRow(['Run Date:', runAt]);
  builder.pushRow(['Operator:', `${operatorName ?? ''}${firmName ? ` (CA firm: ${firmName})` : ''}`]);
  builder.pushRow(['']);
  builder.pushRow(['SUMMARY']);
  builder.beginSummaryBlock();
  builder.pushRow([booksLabel, summary.total_pr_rows]);
  builder.pushRow(['Portal / GSTR Rows', summary.total_portal_rows]);
  builder.pushRow(['IMS Entries', summary.total_ims_rows]);
  builder.pushRow(['Exact Matched', summary.exact_matched, taxMatchedLabel, summary.itc_matched]);
  builder.pushRow(['Matched (fallback, no invoice no.)', summary.matched_fallback]);
  builder.pushRow(['Partial Matched', summary.partial_matched]);
  builder.pushRow(['Credit Notes', summary.credit_notes]);
  builder.pushRow([booksOnlyLabel, summary.pr_only, taxAtRiskLabel, summary.itc_at_risk]);
  builder.pushRow(['  Blank GSTIN in books', summary.mismatch_blank_gstin]);
  builder.pushRow([
    '    ...likely matched to a portal invoice (confirm GSTIN)',
    summary.mismatch_blank_gstin_likely_matched,
  ]);
  builder.pushRow(['  Blank taxable value', summary.mismatch_blank_taxable_value]);
  builder.pushRow(['  Ambiguous rate slab (needs CA review)', summary.mismatch_ambiguous_rate_slab]);
  builder.pushRow(['  Amount mismatch', summary.mismatch_amount]);
  builder.pushRow(['  Date mismatch', summary.mismatch_date]);
  builder.pushRow(['  Genuinely missing', summary.mismatch_genuinely_missing]);
  builder.pushRow([
    'Possible GSTIN mismatch (same vendor, different registration)',
    summary.gstin_mismatch_count,
  ]);
  builder.pushRow(['Portal Only', summary.portal_only]);
  builder.pushRow(['Cross-GSTIN Exceptions', summary.cross_gstin_exception_count ?? 0]);
  builder.pushRow(['IMS Rejected', summary.ims_rejected, 'ITC lost', summary.itc_ims_rejected]);
  builder.pushRow(['IMS Pending', summary.ims_pending, 'ITC deferred', summary.itc_ims_pending]);
  builder.pushRow(['RCM Flags', summary.rcm_flagged, 'RCM payable', summary.rcm_payable]);
  builder.endSummaryBlock();
  builder.pushRow(['']);
  // DETAIL always gets its title + header row, even with zero result rows (unlike the
  // reason-grouped sections in mapMissedBooksToSheetActions, which omit empty sections).
  const detailTitleRow = builder.pushRow(['DETAIL']);
  const detailHeaderRow = builder.pushRow(headers);
  const detailDataStart = builder.rows.length;
  for (const r of tableRows) builder.pushRow(r);
  const detailDataEnd = builder.rows.length - 1;
  builder.sections.push({
    titleRow: detailTitleRow,
    headerRow: detailHeaderRow,
    headers,
    dataRowRange: tableRows.length ? [detailDataStart, detailDataEnd] : null,
  });
  builder.pushRow(['']);
  builder.pushRow(['CELLIX | Reconciliation aid. CA to verify.']);

  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = builder.rows.slice(1);

  const stylingActions = buildReconSheetStylingActions({
    sheetName,
    columnCount: width,
    sheetTitleRow,
    sections: builder.sections,
    summaryValueRange: builder.summaryValueRange,
  });

  const actions: SheetActionPayload[] = [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers: writeHeaders as string[],
      rows: writeRows,
    },
    ...stylingActions,
  ];

  return actions;
}

const BOOKS_ROW_SECTION_HEADERS = [
  'GSTIN',
  'Vendor Name',
  'Invoice No',
  'Invoice Date',
  'Taxable Value',
  'IGST',
  'CGST',
  'SGST',
  'Tax Amount',
  'Document Type',
  'Reason',
  'Explanation',
  'Source Row',
];

const PORTAL_ROW_SECTION_HEADERS = [
  'GSTIN',
  'Vendor Name',
  'Invoice No',
  'Invoice Date',
  'Taxable Value',
  'IGST',
  'CGST',
  'SGST',
  'Tax Amount',
  'Document Type',
  'Source Row',
];

/** Books-side sections source the vendor name from the Purchase Register's Particulars column (mapped to narration). */
function toBooksSectionRow(r: GstResultRow): unknown[] {
  const row = r.registerRow!;
  return [
    row.gstin || '',
    row.narration || '',
    row.invoiceNumber,
    row.invoiceDate,
    row.taxableValue,
    row.igst,
    row.cgst,
    row.sgst,
    row.taxAmount,
    row.documentType,
    r.mismatchReason ?? '',
    r.explanation ?? '',
    refLabel(row.sourceRowRef) ?? '',
  ];
}

/** Portal-only rows have no books-side row — vendor name comes from the portal file's own NAME column (also mapped to narration). */
function toPortalSectionRow(r: GstResultRow): unknown[] {
  const row = r.portalRow!;
  return [
    row.gstin || '',
    row.narration || '',
    row.invoiceNumber,
    row.invoiceDate,
    row.taxableValue,
    row.igst,
    row.cgst,
    row.sgst,
    row.taxAmount,
    row.documentType,
    refLabel(row.sourceRowRef) ?? '',
  ];
}

const GSTIN_MISMATCH_SECTION_HEADERS = [
  'Vendor Name',
  'Books GSTIN',
  'Portal GSTIN',
  'Invoice Date (Books)',
  'Invoice Date (Portal)',
  'Taxable Value',
  'Portal Invoice No',
  'Explanation',
  'Books Source Row',
  'Portal Source Row',
];

/** Both GSTINs shown side by side — the whole point of this section is spotting the registration mismatch at a glance. */
function toGstinMismatchSectionRow(r: GstResultRow): unknown[] {
  const books = r.registerRow!;
  const portal = r.portalRow!;
  return [
    books.narration || portal.narration || '',
    books.gstin || '',
    portal.gstin || '',
    books.invoiceDate,
    portal.invoiceDate,
    books.taxableValue,
    portal.invoiceNumber,
    r.explanation ?? '',
    refLabel(books.sourceRowRef) ?? '',
    refLabel(portal.sourceRowRef) ?? '',
  ];
}

const BLANK_GSTIN_LIKELY_MATCHED_HEADERS = [
  'Vendor Name (Books)',
  'Invoice Date',
  'Taxable Value',
  'Suggested GSTIN',
  'Suggested Vendor Name',
  'Suggested Invoice No',
  'Explanation',
  'Books Source Row',
  'Portal Source Row',
];

/** The suggested GSTIN/vendor/invoice come from the unique portal_only candidate this row was resolved against (closestPortalRow). */
function toBlankGstinLikelyMatchedRow(r: GstResultRow): unknown[] {
  const books = r.registerRow!;
  const portal = r.closestPortalRow;
  return [
    books.narration || '',
    books.invoiceDate,
    books.taxableValue,
    portal?.gstin ?? '',
    portal?.narration ?? '',
    portal?.invoiceNumber ?? '',
    r.explanation ?? '',
    refLabel(books.sourceRowRef) ?? '',
    portal ? (refLabel(portal.sourceRowRef) ?? '') : '',
  ];
}

/**
 * Casual recon: write books rows that did not match the portal, grouped into labeled
 * sections by mismatch reason (never one flat unmatched list), plus portal-only rows.
 */
export function mapMissedBooksToSheetActions(params: {
  sheetName: string;
  portalLabel: string;
  booksLabel: string;
  runAt: string;
  relativeTo?: string;
  resultRows: GstResultRow[];
}): SheetActionPayload[] {
  const { sheetName, portalLabel, booksLabel, runAt, relativeTo, resultRows } = params;
  const missed = resultRows.filter((r) => r.status === 'PR_ONLY' && r.registerRow);
  const portalOnly = resultRows.filter((r) => r.status === 'PORTAL_ONLY' && r.portalRow);
  const gstinMismatchRows = resultRows.filter(
    (r) => r.status === 'GSTIN_MISMATCH' && r.registerRow && r.portalRow,
  );
  if (!missed.length && !portalOnly.length && !gstinMismatchRows.length) return [];

  const byReason = (...reasons: MismatchReason[]) =>
    missed.filter((r) => r.mismatchReason && reasons.includes(r.mismatchReason));

  const blankGstinRows = byReason('blank_counterparty_gstin');
  const blankGstinLikelyMatchedRows = byReason('blank_gstin_likely_matched');
  const blankValueRows = byReason('blank_taxable_value');
  const ambiguousRateSlabRows = byReason('ambiguous_rate_slab');
  const amountMismatchRows = byReason('amount_mismatch');
  const dateMismatchRows = byReason('date_mismatch');
  const genuinelyMissingRows = byReason(...GENUINELY_MISSING_REASONS);
  // Defensive catch-all: a PR_ONLY row must never silently vanish from the sheet even
  // if it somehow reaches here without a recognized reason (e.g. an older/raw fixture).
  const classified = new Set([
    ...blankGstinRows,
    ...blankGstinLikelyMatchedRows,
    ...blankValueRows,
    ...ambiguousRateSlabRows,
    ...amountMismatchRows,
    ...dateMismatchRows,
    ...genuinelyMissingRows,
  ]);
  const unclassifiedRows = missed.filter((r) => !classified.has(r));

  const matchedExact = resultRows.filter((r) => r.status === 'MATCHED' && r.pass === 1).length;
  const matchedFallback = resultRows.filter((r) => r.status === 'MATCHED' && r.pass === 2).length;
  const matchedPartial = resultRows.filter((r) => r.status === 'PARTIAL').length;
  const matchedCreditNote = resultRows.filter((r) => r.status === 'CREDIT_NOTE').length;

  const width = BOOKS_ROW_SECTION_HEADERS.length;
  const title = `${booksLabel} vs ${portalLabel} — categorized`;

  const builder = new StyledSheetBuilder(width);
  const sheetTitleRow = builder.pushRow([title]);
  builder.pushRow([
    `Run: ${runAt}`,
    `Missed (books): ${missed.length}`,
    `Portal-only: ${portalOnly.length}`,
    `GSTIN mismatch: ${gstinMismatchRows.length}`,
  ]);
  builder.pushRow(['']);
  builder.pushRow(['SUMMARY']);
  builder.beginSummaryBlock();
  builder.pushRow(['Matched (exact)', matchedExact]);
  builder.pushRow(['Matched (fallback, no invoice no.)', matchedFallback]);
  builder.pushRow(['Matched (partial/fuzzy)', matchedPartial]);
  builder.pushRow(['Matched (credit/debit note)', matchedCreditNote]);
  builder.pushRow(['']);
  builder.pushRow(['Blank GSTIN in books', blankGstinRows.length]);
  builder.pushRow([
    '  ...likely matched to a portal invoice (confirm GSTIN)',
    blankGstinLikelyMatchedRows.length,
  ]);
  builder.pushRow(['Blank taxable value', blankValueRows.length]);
  builder.pushRow(['Ambiguous rate slab (needs CA review)', ambiguousRateSlabRows.length]);
  builder.pushRow(['Amount mismatch', amountMismatchRows.length]);
  builder.pushRow(['Date mismatch', dateMismatchRows.length]);
  builder.pushRow(['Genuinely missing', genuinelyMissingRows.length]);
  builder.pushRow([
    'Possible GSTIN mismatch (same vendor, different registration)',
    gstinMismatchRows.length,
  ]);
  builder.pushRow(['Portal-only (not in books)', portalOnly.length]);
  if (unclassifiedRows.length) {
    builder.pushRow(['Unclassified (needs review)', unclassifiedRows.length]);
  }
  builder.endSummaryBlock();
  builder.pushRow(['']);

  builder.pushSection('Blank GSTIN rows', BOOKS_ROW_SECTION_HEADERS, blankGstinRows.map(toBooksSectionRow));
  builder.pushSection(
    'Blank GSTIN — likely matched (confirm & fill in GSTIN)',
    BLANK_GSTIN_LIKELY_MATCHED_HEADERS,
    blankGstinLikelyMatchedRows.map(toBlankGstinLikelyMatchedRow),
  );
  builder.pushSection('Blank taxable value rows', BOOKS_ROW_SECTION_HEADERS, blankValueRows.map(toBooksSectionRow));
  builder.pushSection(
    'Ambiguous rate slab rows (needs CA review)',
    BOOKS_ROW_SECTION_HEADERS,
    ambiguousRateSlabRows.map(toBooksSectionRow),
  );
  builder.pushSection('Amount mismatch rows', BOOKS_ROW_SECTION_HEADERS, amountMismatchRows.map(toBooksSectionRow));
  builder.pushSection('Date mismatch rows', BOOKS_ROW_SECTION_HEADERS, dateMismatchRows.map(toBooksSectionRow));
  builder.pushSection(
    'Genuinely missing rows',
    BOOKS_ROW_SECTION_HEADERS,
    genuinelyMissingRows.map(toBooksSectionRow),
  );
  builder.pushSection(
    'Possible GSTIN mismatch (same vendor, different registration)',
    GSTIN_MISMATCH_SECTION_HEADERS,
    gstinMismatchRows.map(toGstinMismatchSectionRow),
  );
  builder.pushSection(
    'Unclassified rows (needs review)',
    BOOKS_ROW_SECTION_HEADERS,
    unclassifiedRows.map(toBooksSectionRow),
  );
  builder.pushSection(
    `Portal-only rows (in ${portalLabel}, not in books)`,
    PORTAL_ROW_SECTION_HEADERS,
    portalOnly.map(toPortalSectionRow),
  );
  builder.pushRow(['CELLIX | Reconciliation aid. CA to verify.']);

  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = builder.rows.slice(1);

  const stylingActions = buildReconSheetStylingActions({
    sheetName,
    columnCount: width,
    sheetTitleRow,
    sections: builder.sections,
    summaryValueRange: builder.summaryValueRange,
  });

  return [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers: writeHeaders as string[],
      rows: writeRows,
    },
    ...stylingActions,
  ];
}

/**
 * Books-only flat layout — "row exists in the books, but that same invoice does NOT
 * exist in the portal." Single flat table, no reason-based sections: the union of every
 * books-side reason (blank GSTIN, blank taxable value, ambiguous rate slab, amount
 * mismatch, date mismatch, genuinely missing — PR_ONLY, any reason) plus GSTIN-mismatch
 * rows (same vendor, different registration — still a books-side row with no portal
 * counterpart under its own GSTIN). Reuses the same styling helper as the categorized
 * layout. This is the reverse direction of mapPortalOnlyFlatToSheetActions — never
 * confuse the two: this one never includes a PORTAL_ONLY row.
 */
export function mapMissedBooksFlatToSheetActions(params: {
  sheetName: string;
  portalLabel: string;
  booksLabel: string;
  runAt: string;
  relativeTo?: string;
  resultRows: GstResultRow[];
}): SheetActionPayload[] {
  const { sheetName, portalLabel, booksLabel, runAt, relativeTo, resultRows } = params;
  const missed = resultRows.filter(
    (r) => (r.status === 'PR_ONLY' || r.status === 'GSTIN_MISMATCH') && r.registerRow,
  );
  if (!missed.length) return [];

  const width = BOOKS_ROW_SECTION_HEADERS.length;
  const title = `${booksLabel} rows missing from ${portalLabel}`;

  const builder = new StyledSheetBuilder(width);
  const sheetTitleRow = builder.pushRow([title]);
  builder.pushRow([`Run: ${runAt}`, `Rows: ${missed.length}`]);
  builder.pushRow(['']);
  builder.pushSection('Missing Rows', BOOKS_ROW_SECTION_HEADERS, missed.map(toBooksSectionRow));
  builder.pushRow(['CELLIX | Reconciliation aid. CA to verify.']);

  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = builder.rows.slice(1);

  const stylingActions = buildReconSheetStylingActions({
    sheetName,
    columnCount: width,
    sheetTitleRow,
    sections: builder.sections,
  });

  return [
    { type: 'CREATE_SHEET', sheetName, name: sheetName, relativeTo, position: 'after' },
    { type: 'WRITE_TABLE', sheetName, headers: writeHeaders as string[], rows: writeRows },
    ...stylingActions,
  ];
}

/**
 * Portal-only flat layout — the reverse of mapMissedBooksFlatToSheetActions: "row exists
 * on the portal, but that same invoice does NOT exist in the books." Single flat table,
 * PORTAL_ONLY rows only — never a books-side (PR_ONLY/GSTIN_MISMATCH) row.
 */
export function mapPortalOnlyFlatToSheetActions(params: {
  sheetName: string;
  portalLabel: string;
  booksLabel: string;
  runAt: string;
  relativeTo?: string;
  resultRows: GstResultRow[];
}): SheetActionPayload[] {
  const { sheetName, portalLabel, booksLabel, runAt, relativeTo, resultRows } = params;
  const portalOnly = resultRows.filter((r) => r.status === 'PORTAL_ONLY' && r.portalRow);
  if (!portalOnly.length) return [];

  const width = PORTAL_ROW_SECTION_HEADERS.length;
  const title = `${portalLabel} rows missing from ${booksLabel}`;

  const builder = new StyledSheetBuilder(width);
  const sheetTitleRow = builder.pushRow([title]);
  builder.pushRow([`Run: ${runAt}`, `Rows: ${portalOnly.length}`]);
  builder.pushRow(['']);
  builder.pushSection('Missing Rows', PORTAL_ROW_SECTION_HEADERS, portalOnly.map(toPortalSectionRow));
  builder.pushRow(['CELLIX | Reconciliation aid. CA to verify.']);

  const writeHeaders = builder.rows[0].map((c, i) => (i === 0 ? String(c) : `C${i + 1}`));
  const writeRows = builder.rows.slice(1);

  const stylingActions = buildReconSheetStylingActions({
    sheetName,
    columnCount: width,
    sheetTitleRow,
    sections: builder.sections,
  });

  return [
    { type: 'CREATE_SHEET', sheetName, name: sheetName, relativeTo, position: 'after' },
    { type: 'WRITE_TABLE', sheetName, headers: writeHeaders as string[], rows: writeRows },
    ...stylingActions,
  ];
}

export function map3bVs2bToActions(params: {
  sheetName: string;
  period?: string;
  gstin?: string;
  result: Gstr3bVs2bOutput;
  runAt: string;
  relativeTo?: string;
}): SheetActionPayload[] {
  const { sheetName, period, gstin, result, runAt, relativeTo } = params;
  const headers = ['Metric', 'IGST', 'CGST', 'SGST', 'Total', 'Notes'];
  const rows: unknown[][] = [
    [
      'Variance (3B − 2B)',
      result.variances.igst,
      result.variances.cgst,
      result.variances.sgst,
      result.variances.total,
      result.status,
    ],
    ['Excess claim', '', '', '', result.excessClaim, result.drc01cRisk ? 'DRC-01C RISK' : ''],
    ['Short claim', '', '', '', result.shortClaim, ''],
    ...result.messages.map((m) => ['Note', '', '', '', '', m]),
    ['', '', '', '', '', 'CELLIX | Reconciliation aid. CA to verify.'],
  ];

  return [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      relativeTo,
      position: 'after',
    },
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers,
      rows: [
        [`Period: ${period ?? ''} | GSTIN: ${gstin ?? ''} | Run: ${runAt}`, '', '', '', '', ''],
        ...rows,
      ],
    },
  ];
}
