import {
  DomainException,
  GstMatchMode,
  GstMatchSettings,
  GstReconStatus,
  MatchedPair,
  MismatchReason,
  NormalizedInvoiceRow,
  SourceRef,
} from '../types/domain-tool.types';
import {
  amountWithinTolerance,
  daysBetween,
  isValidGstinFormat,
  stringSimilarityPercent,
} from './normalize-invoice';

export interface GstResultRow {
  status: GstReconStatus;
  pass: number | null;
  confidence: number;
  difference?: string;
  diffType?: string;
  itcAmount: number;
  rcmFlag: boolean;
  imsStatus?: string | null;
  registerRow?: NormalizedInvoiceRow;
  portalRow?: NormalizedInvoiceRow;
  imsRow?: NormalizedInvoiceRow;
  /** Set on PR_ONLY rows — the specific, CA-readable reason this row didn't match. */
  mismatchReason?: MismatchReason;
  explanation?: string;
  /** Best near-match found on the other side, when the reason is amount/date mismatch. */
  closestPortalRow?: NormalizedInvoiceRow;
  fieldDiff?: Array<{ field: string; booksValue: unknown; portalValue: unknown }>;
}

export interface MatchWorkingSet {
  unmatchedPr: NormalizedInvoiceRow[];
  unmatchedPortal: NormalizedInvoiceRow[];
  /** Full original portal set (never mutated) — used to diagnose why an unmatched row didn't match. */
  allPortal: NormalizedInvoiceRow[];
  results: GstResultRow[];
  exceptions: DomainException[];
  mode: GstMatchMode;
}

function rowNum(ref: SourceRef): number {
  return typeof ref.rowOrLine === 'number' ? ref.rowOrLine : Number(ref.rowOrLine) || 0;
}

function itcOf(row: NormalizedInvoiceRow): number {
  return Math.abs(row.taxAmount || row.igst + row.cgst + row.sgst);
}

function isB2c(row: NormalizedInvoiceRow, mode: GstMatchMode): boolean {
  if (mode !== 'sales') return false;
  if (row.supplyCategory === 'b2c') return true;
  return !row.gstin;
}

export function createWorkingSet(
  pr: NormalizedInvoiceRow[],
  portal: NormalizedInvoiceRow[],
  mode: GstMatchMode = 'purchase',
): MatchWorkingSet {
  return {
    unmatchedPr: [...pr],
    unmatchedPortal: [...portal],
    allPortal: [...portal],
    results: [],
    exceptions: [],
    mode,
  };
}

function exactKeys(row: NormalizedInvoiceRow, mode: GstMatchMode): string[] {
  const keys: string[] = [];
  if (isB2c(row, mode) && row.normalizedInvoiceNumber) {
    keys.push(`B2C|${row.normalizedInvoiceNumber}`);
  } else if (row.gstin && row.normalizedInvoiceNumber) {
    keys.push(`${row.gstin}|${row.normalizedInvoiceNumber}`);
  }
  if (row.irn) {
    keys.push(`IRN|${row.irn.toUpperCase()}`);
  }
  return keys;
}

/** Pass 1: exact GSTIN + normalized invoice + amount ±0.01 (or IRN / B2C bucket). */
export function runPassExact(ws: MatchWorkingSet): void {
  const portalByKey = new Map<string, number[]>();
  for (let i = 0; i < ws.unmatchedPortal.length; i++) {
    const p = ws.unmatchedPortal[i];
    const keys = exactKeys(p, ws.mode);
    for (const k of keys) {
      const list = portalByKey.get(k) ?? [];
      list.push(i);
      portalByKey.set(k, list);
    }
  }

  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const keys = exactKeys(pr, ws.mode);
    let matchedIdx = -1;
    for (const k of keys) {
      const candidates = portalByKey.get(k) ?? [];
      for (const idx of candidates) {
        if (usedPortal.has(idx)) continue;
        const portal = ws.unmatchedPortal[idx];
        if (
          amountWithinTolerance(pr.taxableValue ?? 0, portal.taxableValue ?? 0, 0.01, 0) ||
          amountWithinTolerance(pr.taxAmount, portal.taxAmount, 0.01, 0)
        ) {
          matchedIdx = idx;
          break;
        }
      }
      if (matchedIdx >= 0) break;
    }

    if (matchedIdx < 0) {
      stillPr.push(pr);
      continue;
    }

    usedPortal.add(matchedIdx);
    const portal = ws.unmatchedPortal[matchedIdx];
    ws.results.push({
      status: 'MATCHED',
      pass: 1,
      confidence: 1,
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
    });
  }

  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

/**
 * Pass 1b: fallback match on GSTIN + exact invoice date + taxable value (within tolerance).
 * Runs ONLY when the books sheet has no Invoice Number column at all — the scenario that
 * previously produced 0 matched / everything dumped into PR_ONLY. Checked once per sheet
 * by the caller (gstMatch), not per row.
 */
export function runPass1bFallback(ws: MatchWorkingSet, settings: GstMatchSettings): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  const portalByGstinDate = new Map<string, number[]>();
  for (let i = 0; i < ws.unmatchedPortal.length; i++) {
    const p = ws.unmatchedPortal[i];
    if (!p.gstin || !p.invoiceDate) continue;
    const key = `${p.gstin}|${p.invoiceDate}`;
    const list = portalByGstinDate.get(key) ?? [];
    list.push(i);
    portalByGstinDate.set(key, list);
  }

  for (const pr of ws.unmatchedPr) {
    if (!pr.gstin || !pr.invoiceDate || pr.taxableValue == null) {
      stillPr.push(pr);
      continue;
    }
    const candidates = portalByGstinDate.get(`${pr.gstin}|${pr.invoiceDate}`) ?? [];
    const matchedIdx = candidates.find(
      (idx) =>
        !usedPortal.has(idx) &&
        ws.unmatchedPortal[idx].taxableValue != null &&
        amountWithinTolerance(
          pr.taxableValue as number,
          ws.unmatchedPortal[idx].taxableValue as number,
          settings.amountToleranceAbs,
          settings.amountTolerancePct,
        ),
    );

    if (matchedIdx === undefined) {
      stillPr.push(pr);
      continue;
    }

    usedPortal.add(matchedIdx);
    const portal = ws.unmatchedPortal[matchedIdx];
    ws.results.push({
      status: 'MATCHED',
      pass: 2,
      confidence: 0.95,
      difference: 'Matched on GSTIN + Date + Amount (no invoice number available in source register).',
      diffType: 'FALLBACK_NO_INVOICE_NUMBER',
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
    });
  }

  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

/** Pass 3: same GSTIN (or B2C bucket), fuzzy invoice, amount/date tolerances. */
export function runPassFuzzy(ws: MatchWorkingSet, settings: GstMatchSettings): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const prB2c = isB2c(pr, ws.mode);
    if (!pr.gstin && !prB2c) {
      stillPr.push(pr);
      continue;
    }
    let bestIdx = -1;
    let bestScore = 0;
    let bestDiff = '';

    for (let i = 0; i < ws.unmatchedPortal.length; i++) {
      if (usedPortal.has(i)) continue;
      const portal = ws.unmatchedPortal[i];
      const portalB2c = isB2c(portal, ws.mode);
      if (prB2c) {
        if (!portalB2c) continue;
      } else if (portal.gstin !== pr.gstin) {
        continue;
      }

      const sim = stringSimilarityPercent(
        pr.normalizedInvoiceNumber,
        portal.normalizedInvoiceNumber,
      );
      if (sim < settings.invoiceFuzzyThreshold) continue;
      if (
        !amountWithinTolerance(
          pr.taxableValue ?? 0,
          portal.taxableValue ?? 0,
          settings.amountToleranceAbs,
          settings.amountTolerancePct,
        ) &&
        !amountWithinTolerance(
          pr.taxAmount,
          portal.taxAmount,
          settings.amountToleranceAbs,
          settings.amountTolerancePct,
        )
      ) {
        continue;
      }
      const dayDiff = daysBetween(pr.invoiceDate, portal.invoiceDate);
      if (
        dayDiff !== null &&
        settings.dateToleranceDays >= 0 &&
        dayDiff > settings.dateToleranceDays
      ) {
        continue;
      }
      if (sim > bestScore) {
        bestScore = sim;
        bestIdx = i;
        bestDiff = `${pr.invoiceNumber} vs ${portal.invoiceNumber}`;
      }
    }

    if (bestIdx < 0) {
      stillPr.push(pr);
      continue;
    }

    usedPortal.add(bestIdx);
    const portal = ws.unmatchedPortal[bestIdx];
    const conf = Math.min(0.99, bestScore / 100);
    const amtDiff = Math.abs((pr.taxableValue ?? 0) - (portal.taxableValue ?? 0));
    let diffType = 'INVOICE_FORMAT';
    let difference = bestDiff;
    if (amtDiff > 0.01) {
      diffType = 'AMOUNT_VARIANCE';
      difference = `PR ${pr.taxableValue} vs portal ${portal.taxableValue} (${bestDiff})`;
    } else {
      const d = daysBetween(pr.invoiceDate, portal.invoiceDate);
      if (d !== null && d > 0) {
        diffType = 'DATE_VARIANCE';
        difference = `dates differ by ${d}d; ${bestDiff}`;
      }
    }

    ws.results.push({
      status: 'PARTIAL',
      pass: 3,
      confidence: conf,
      difference,
      diffType,
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
    });
    ws.exceptions.push({
      code: 'GST_FUZZY_INVOICE_MATCH',
      severity: 'flag',
      message: `Partial match: ${difference}`,
      affectedRows: [rowNum(pr.sourceRowRef), rowNum(portal.sourceRowRef)],
    });
  }

  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

/** Pass 4: credit/debit notes by GSTIN + absolute amount proximity + 90 day window. */
export function runPassCdn(ws: MatchWorkingSet): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const isCdn =
      pr.documentType === 'credit_note' ||
      pr.documentType === 'debit_note' ||
      (pr.taxableValue ?? 0) < 0 ||
      /credit|cdn|debit/i.test(pr.narration);

    if (!isCdn || !pr.gstin) {
      stillPr.push(pr);
      continue;
    }

    let bestIdx = -1;
    let bestGap = Number.POSITIVE_INFINITY;

    for (let i = 0; i < ws.unmatchedPortal.length; i++) {
      if (usedPortal.has(i)) continue;
      const portal = ws.unmatchedPortal[i];
      if (portal.gstin !== pr.gstin) continue;
      const portalCdn =
        portal.documentType === 'credit_note' ||
        portal.documentType === 'debit_note' ||
        (portal.taxableValue ?? 0) < 0 ||
        /credit|cdn|debit/i.test(portal.narration);
      if (!portalCdn && portal.documentType !== 'amended') continue;

      const gap = Math.abs(Math.abs(pr.taxableValue ?? 0) - Math.abs(portal.taxableValue ?? 0));
      const dayDiff = daysBetween(pr.invoiceDate, portal.invoiceDate);
      if (dayDiff !== null && dayDiff > 90) continue;
      if (gap < bestGap) {
        bestGap = gap;
        bestIdx = i;
      }
    }

    if (bestIdx < 0 || bestGap > Math.max(1, Math.abs(pr.taxableValue ?? 0) * 0.01)) {
      stillPr.push(pr);
      continue;
    }

    usedPortal.add(bestIdx);
    const portal = ws.unmatchedPortal[bestIdx];
    ws.results.push({
      status: 'CREDIT_NOTE',
      pass: 4,
      confidence: bestGap <= 0.01 ? 0.98 : 0.9,
      difference: `CDN match: ${pr.invoiceNumber} ↔ ${portal.invoiceNumber}`,
      diffType: 'CREDIT_NOTE',
      itcAmount: -itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
    });
  }

  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

const RCM_KEYWORDS =
  /\b(freight|gta|goods transport|advocate|legal fee|legal service|google ads|google|aws|amazon web|meta ads|facebook ads|import of service|oidar|security service|rent.*unregistered)\b/i;

/** Pass 5: RCM detection on remaining PR rows. */
export function runPassRcm(ws: MatchWorkingSet, settings: GstMatchSettings): void {
  if (!settings.detectRcm) return;
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const blankGstin = !pr.gstin || !isValidGstinFormat(pr.gstin);
    const keywordHit = RCM_KEYWORDS.test(pr.narration);
    if (blankGstin || keywordHit) {
      ws.results.push({
        status: 'RCM',
        pass: 5,
        confidence: blankGstin && keywordHit ? 0.85 : 0.7,
        difference: blankGstin
          ? 'Blank/invalid GSTIN — potential URD/RCM'
          : `RCM keyword in narration: ${pr.narration.slice(0, 80)}`,
        diffType: 'RCM',
        itcAmount: 0,
        rcmFlag: true,
        registerRow: pr,
      });
      ws.exceptions.push({
        code: 'GST_RCM_FLAG',
        severity: 'flag',
        message: `RCM flagged for PR row ${rowNum(pr.sourceRowRef)}`,
        affectedRows: [rowNum(pr.sourceRowRef)],
      });
    } else {
      stillPr.push(pr);
    }
  }

  ws.unmatchedPr = stillPr;
}

/**
 * Pass 6: IMS action cross-check against remaining + matched results.
 * IMS rows may be provided separately; attach status by GSTIN+invoice key.
 */
export function runPassIms(
  ws: MatchWorkingSet,
  imsRows: NormalizedInvoiceRow[],
  settings: GstMatchSettings,
): void {
  if (!settings.useImsData || !imsRows.length) return;

  const imsByKey = new Map<string, NormalizedInvoiceRow>();
  for (const row of imsRows) {
    if (row.gstin && row.normalizedInvoiceNumber) {
      imsByKey.set(`${row.gstin}|${row.normalizedInvoiceNumber}`, row);
    }
  }

  for (const result of ws.results) {
    const keyRow = result.registerRow ?? result.portalRow;
    if (!keyRow) continue;
    const ims = imsByKey.get(`${keyRow.gstin}|${keyRow.normalizedInvoiceNumber}`);
    if (!ims) continue;
    result.imsRow = ims;
    result.imsStatus = ims.imsAction ?? null;
    imsByKey.delete(`${keyRow.gstin}|${keyRow.normalizedInvoiceNumber}`);

    if (ims.imsAction === 'Rejected') {
      result.status = 'IMS_REJECTED';
      result.difference = (result.difference ? result.difference + '; ' : '') + 'IMS Rejected — ITC not claimable';
      result.diffType = 'IMS_REJECTED';
      result.pass = result.pass ?? 6;
      ws.exceptions.push({
        code: 'GST_IMS_REJECTED',
        severity: 'flag',
        message: `IMS rejected invoice ${keyRow.invoiceNumber}`,
        affectedRows: [rowNum(keyRow.sourceRowRef)],
      });
    } else if (ims.imsAction === 'Pending') {
      if (result.status === 'MATCHED' || result.status === 'PARTIAL') {
        result.status = 'IMS_PENDING';
      }
      result.difference =
        (result.difference ? result.difference + '; ' : '') + 'IMS Pending — ITC deferred';
      result.diffType = 'IMS_PENDING';
      result.pass = 6;
    } else if (ims.imsAction === 'AutoAccepted') {
      result.difference =
        (result.difference ? result.difference + '; ' : '') + 'IMS Auto-accepted — review recommended';
      result.diffType = result.diffType ?? 'IMS_AUTO_ACCEPT';
      if (result.status === 'MATCHED') {
        // keep MATCHED but flag
        ws.exceptions.push({
          code: 'GST_IMS_AUTO_ACCEPT',
          severity: 'flag',
          message: `IMS auto-accepted without review: ${keyRow.invoiceNumber}`,
          affectedRows: [rowNum(keyRow.sourceRowRef)],
        });
      } else {
        result.status = 'IMS_AUTO_ACCEPT';
        result.pass = 6;
      }
    }
  }

  // PR-only later will still be unmatched; attach IMS for PR keys still available
  // Remaining IMS without PR/portals becomes IMS_ONLY
  const stillIms = [...imsByKey.values()];
  const usedIms = new Set<string>();

  for (const pr of ws.unmatchedPr) {
    const key = `${pr.gstin}|${pr.normalizedInvoiceNumber}`;
    const ims = imsByKey.get(key);
    if (!ims) continue;
    usedIms.add(key);
    if (ims.imsAction === 'Rejected') {
      ws.results.push({
        status: 'IMS_REJECTED',
        pass: 6,
        confidence: 0.95,
        difference: 'In PR and Rejected in IMS — ITC lost',
        diffType: 'IMS_REJECTED',
        itcAmount: itcOf(pr),
        rcmFlag: false,
        registerRow: pr,
        imsRow: ims,
        imsStatus: ims.imsAction,
      });
    } else if (ims.imsAction === 'Pending') {
      ws.results.push({
        status: 'IMS_PENDING',
        pass: 6,
        confidence: 0.9,
        difference: 'In PR, Pending in IMS — ITC deferred',
        diffType: 'IMS_PENDING',
        itcAmount: itcOf(pr),
        rcmFlag: false,
        registerRow: pr,
        imsRow: ims,
        imsStatus: ims.imsAction,
      });
    }
  }
  ws.unmatchedPr = ws.unmatchedPr.filter(
    (pr) => !usedIms.has(`${pr.gstin}|${pr.normalizedInvoiceNumber}`),
  );

  for (const ims of stillIms) {
    const key = `${ims.gstin}|${ims.normalizedInvoiceNumber}`;
    if (usedIms.has(key)) continue;
    // only if not linked already
    if (ws.results.some((r) => r.imsRow === ims)) continue;
    const onPortal = ws.unmatchedPortal.find(
      (p) => p.gstin === ims.gstin && p.normalizedInvoiceNumber === ims.normalizedInvoiceNumber,
    );
    if (onPortal) continue;
    ws.results.push({
      status: 'IMS_ONLY',
      pass: 6,
      confidence: 0.9,
      difference: 'In IMS only — update books if valid',
      diffType: 'IMS_ONLY',
      itcAmount: itcOf(ims),
      rcmFlag: false,
      imsRow: ims,
      imsStatus: ims.imsAction ?? null,
    });
  }
}

interface Diagnosis {
  reason: MismatchReason;
  explanation: string;
  closestPortalRow?: NormalizedInvoiceRow;
  fieldDiff?: Array<{ field: string; booksValue: unknown; portalValue: unknown }>;
}

function closestByAmount(candidates: NormalizedInvoiceRow[], target: number): NormalizedInvoiceRow {
  return candidates.reduce((best, c) =>
    Math.abs((c.taxableValue ?? 0) - target) < Math.abs((best.taxableValue ?? 0) - target) ? c : best,
  );
}

function closestByDate(candidates: NormalizedInvoiceRow[], targetDate: string): NormalizedInvoiceRow {
  const targetTime = Date.parse(targetDate) || 0;
  return candidates.reduce((best, c) => {
    const cDiff = Math.abs((Date.parse(c.invoiceDate) || 0) - targetTime);
    const bestDiff = Math.abs((Date.parse(best.invoiceDate) || 0) - targetTime);
    return cDiff < bestDiff ? c : best;
  });
}

/**
 * Every books row that survives Pass 1/1b/2/3 unmatched gets a specific, CA-readable
 * reason instead of a flat "unmatched" tag. `otherSideForGstin` should be the full
 * (not just still-unmatched) set of portal rows sharing this row's GSTIN, so a genuinely
 * missing GSTIN can be told apart from one whose invoices were all already matched elsewhere.
 */
export function diagnoseUnmatchedRow(
  row: NormalizedInvoiceRow,
  otherSideForGstin: NormalizedInvoiceRow[],
): Diagnosis {
  if (row.ambiguousRateSlab) {
    return {
      reason: 'ambiguous_rate_slab',
      explanation:
        row.ambiguousRateSlabDetail ??
        'Multiple rate-slab columns are populated for this row and the correct one could not be determined automatically — needs CA review.',
    };
  }
  if (!row.gstin) {
    return {
      reason: 'blank_counterparty_gstin',
      explanation: 'GSTIN is blank in the register for this row — cannot be matched.',
    };
  }
  if (row.taxableValue == null) {
    return {
      reason: 'blank_taxable_value',
      explanation: 'No taxable value found in any rate column for this row.',
    };
  }
  if (otherSideForGstin.length === 0) {
    return {
      reason: 'gstin_not_in_portal',
      explanation: `GSTIN ${row.gstin} does not appear anywhere in the portal file for this period.`,
    };
  }

  const sameDate = otherSideForGstin.filter((p) => p.invoiceDate === row.invoiceDate);
  if (sameDate.length > 0) {
    const closest = closestByAmount(sameDate, row.taxableValue);
    return {
      reason: 'amount_mismatch',
      closestPortalRow: closest,
      fieldDiff: [{ field: 'taxableValue', booksValue: row.taxableValue, portalValue: closest.taxableValue }],
      explanation: `Same GSTIN and date found in portal, but amount differs (books ₹${row.taxableValue} vs portal ₹${closest.taxableValue}).`,
    };
  }

  const sameAmount = otherSideForGstin.filter(
    (p) => p.taxableValue != null && amountWithinTolerance(p.taxableValue, row.taxableValue as number, 1.0, 0.5),
  );
  if (sameAmount.length > 0) {
    const closest = closestByDate(sameAmount, row.invoiceDate);
    return {
      reason: 'date_mismatch',
      closestPortalRow: closest,
      fieldDiff: [{ field: 'invoiceDate', booksValue: row.invoiceDate, portalValue: closest.invoiceDate }],
      explanation: `Same GSTIN and amount found in portal, but date differs (books ${row.invoiceDate} vs portal ${closest.invoiceDate}).`,
    };
  }

  return {
    reason: 'genuinely_missing',
    explanation: 'No matching invoice found in the portal for this GSTIN, date, or amount — likely not filed by the supplier.',
  };
}

/** GSTIN characters 3-12 (1-based) are the PAN embedded in every Indian GSTIN. */
function extractPan(gstin: string): string | null {
  return gstin.length === 15 ? gstin.slice(2, 12) : null;
}

function buildGstinMismatchExplanation(
  books: NormalizedInvoiceRow,
  portal: NormalizedInvoiceRow,
  pan: string,
): string {
  const sameDate = books.invoiceDate === portal.invoiceDate;
  const dateClause = sameDate
    ? 'same date and amount'
    : `same amount (but the invoice date also differs — books ${books.invoiceDate} vs portal ${portal.invoiceDate})`;
  return (
    `Same vendor (PAN ${pan}), ${dateClause}, but booked under GSTIN ${books.gstin} in your register ` +
    `vs GSTIN ${portal.gstin} in GSTR-2B — check which registration this vendor actually used for this invoice.`
  );
}

/**
 * Catches the case where the same vendor (same PAN — GSTIN characters 3-12) was booked
 * under a different GSTIN registration than the one it filed under, e.g. a single wrong
 * digit in the state code or checksum. Looks like a genuinely missing invoice on both
 * sides, but is really a registration mismatch, not a missing filing.
 *
 * Runs ONLY on rows that survive every earlier pass (exact, Pass 1b fallback, fuzzy,
 * credit/debit note, RCM) unmatched — called at the start of finalizeUnmatched, so it
 * never interferes with normal GSTIN-exact matching. Amount must match within the usual
 * tolerance; date is preferred but not required (real books-vs-portal date entry gaps do
 * happen for the same invoice, and requiring an exact date would miss real registration
 * mismatches) — when dates differ, the explanation says so explicitly rather than
 * implying an exact match.
 */
export function runPanCrossGstinMatch(ws: MatchWorkingSet): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  const portalIdxByPan = new Map<string, number[]>();
  for (let i = 0; i < ws.unmatchedPortal.length; i++) {
    const p = ws.unmatchedPortal[i];
    const pan = p.gstin ? extractPan(p.gstin) : null;
    if (!pan) continue;
    const list = portalIdxByPan.get(pan) ?? [];
    list.push(i);
    portalIdxByPan.set(pan, list);
  }

  for (const pr of ws.unmatchedPr) {
    const pan = pr.gstin ? extractPan(pr.gstin) : null;
    if (!pan || pr.taxableValue == null) {
      stillPr.push(pr);
      continue;
    }

    const candidates = (portalIdxByPan.get(pan) ?? []).filter(
      (idx) => !usedPortal.has(idx) && ws.unmatchedPortal[idx].gstin !== pr.gstin,
    );
    const amountMatched = candidates.filter((idx) => {
      const p = ws.unmatchedPortal[idx];
      return p.taxableValue != null && amountWithinTolerance(pr.taxableValue as number, p.taxableValue, 1.0, 0.5);
    });

    if (!amountMatched.length) {
      stillPr.push(pr);
      continue;
    }

    const exactDateIdx = amountMatched.find((idx) => ws.unmatchedPortal[idx].invoiceDate === pr.invoiceDate);
    const chosenIdx =
      exactDateIdx !== undefined
        ? exactDateIdx
        : amountMatched.reduce((bestIdx, idx) => {
            const target = Date.parse(pr.invoiceDate) || 0;
            const bestDiff = Math.abs((Date.parse(ws.unmatchedPortal[bestIdx].invoiceDate) || 0) - target);
            const curDiff = Math.abs((Date.parse(ws.unmatchedPortal[idx].invoiceDate) || 0) - target);
            return curDiff < bestDiff ? idx : bestIdx;
          }, amountMatched[0]);

    usedPortal.add(chosenIdx);
    const portal = ws.unmatchedPortal[chosenIdx];
    const explanation = buildGstinMismatchExplanation(pr, portal, pan);
    ws.results.push({
      status: 'GSTIN_MISMATCH',
      pass: null,
      confidence: 0.85,
      difference: explanation,
      diffType: 'GSTIN_MISMATCH_SAME_PAN',
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      mismatchReason: 'gstin_mismatch_same_pan',
      explanation,
      imsStatus: portal.imsAction ?? null,
    });
  }

  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

/** Finalize unmatched PR / portal rows — every PR_ONLY row is diagnosed with a specific reason. */
export function finalizeUnmatched(ws: MatchWorkingSet): void {
  runPanCrossGstinMatch(ws);
  for (const pr of ws.unmatchedPr) {
    const candidatesForGstin = ws.allPortal.filter((p) => p.gstin && p.gstin === pr.gstin);
    const diagnosis = diagnoseUnmatchedRow(pr, candidatesForGstin);
    ws.results.push({
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      difference: diagnosis.explanation,
      diffType: diagnosis.reason.toUpperCase(),
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
      mismatchReason: diagnosis.reason,
      explanation: diagnosis.explanation,
      closestPortalRow: diagnosis.closestPortalRow,
      fieldDiff: diagnosis.fieldDiff,
    });
  }
  for (const portal of ws.unmatchedPortal) {
    ws.results.push({
      status: 'PORTAL_ONLY',
      pass: null,
      confidence: 1,
      difference: 'On portal only — not found in purchase register',
      diffType: 'PORTAL_ONLY',
      itcAmount: itcOf(portal),
      rcmFlag: false,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
    });
  }
  ws.unmatchedPr = [];
  ws.unmatchedPortal = [];

  resolveLikelyBlankGstinMatches(ws);
}

/**
 * A books row with a blank GSTIN never gets compared against portal data (GSTIN is the
 * primary match key) — but its real-world counterpart invoice still exists in the portal
 * file and lands in portal_only, so the same discrepancy gets counted and displayed
 * twice: once as blank_counterparty_gstin, once as portal_only.
 *
 * Runs once, after PR_ONLY/PORTAL_ONLY are both fully populated: for each
 * blank_counterparty_gstin row, look for a UNIQUE portal_only candidate matching on
 * invoice date + taxable value (independent of GSTIN, since the books side has none).
 * Exactly one candidate → reclassify the books row as blank_gstin_likely_matched with
 * the candidate's GSTIN/vendor/invoice attached as a suggestion, and drop that portal
 * row from portal_only entirely (it's the same real-world invoice, not two). Zero or
 * multiple candidates → leave both sides unchanged; never guess.
 */
export function resolveLikelyBlankGstinMatches(ws: MatchWorkingSet): void {
  const portalOnlyResults = ws.results.filter((r) => r.status === 'PORTAL_ONLY' && r.portalRow);
  if (!portalOnlyResults.length) return;

  const usedPortalResults = new Set<GstResultRow>();

  for (const r of ws.results) {
    if (r.mismatchReason !== 'blank_counterparty_gstin' || !r.registerRow) continue;
    const row = r.registerRow;
    if (row.taxableValue == null) continue;

    const candidates = portalOnlyResults.filter((p) => {
      if (usedPortalResults.has(p)) return false;
      const portal = p.portalRow!;
      return (
        portal.invoiceDate === row.invoiceDate &&
        portal.taxableValue != null &&
        amountWithinTolerance(row.taxableValue as number, portal.taxableValue, 1.0, 0.5)
      );
    });

    if (candidates.length !== 1) continue;

    const match = candidates[0];
    usedPortalResults.add(match);
    const portal = match.portalRow!;
    const explanation =
      `GSTIN is blank in the register, but a portal invoice from ${portal.narration || 'this vendor'} ` +
      `(GSTIN ${portal.gstin}) matches this row's date and amount — likely the correct vendor. Confirm and fill in the GSTIN.`;
    r.mismatchReason = 'blank_gstin_likely_matched';
    r.explanation = explanation;
    r.difference = explanation;
    r.diffType = 'BLANK_GSTIN_LIKELY_MATCHED';
    r.closestPortalRow = portal;
  }

  if (usedPortalResults.size) {
    ws.results = ws.results.filter((r) => !usedPortalResults.has(r));
  }
}

export function toMatchedPairs(results: GstResultRow[]): MatchedPair[] {
  return results
    .filter((r) => r.registerRow && r.portalRow && (r.status === 'MATCHED' || r.status === 'PARTIAL' || r.status === 'CREDIT_NOTE'))
    .map((r) => ({
      registerRow: r.registerRow!,
      portalRow: r.portalRow!,
      matchKeys: ['gstin', 'invoiceNumber'],
      confidence: r.confidence,
      pass: r.pass ?? 0,
      difference: r.difference,
      diffType: r.diffType,
    }));
}
