import {
  DomainException,
  GstMatchSettings,
  GstReconStatus,
  MatchedPair,
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
}

export interface MatchWorkingSet {
  unmatchedPr: NormalizedInvoiceRow[];
  unmatchedPortal: NormalizedInvoiceRow[];
  results: GstResultRow[];
  exceptions: DomainException[];
}

function rowNum(ref: SourceRef): number {
  return typeof ref.rowOrLine === 'number' ? ref.rowOrLine : Number(ref.rowOrLine) || 0;
}

function itcOf(row: NormalizedInvoiceRow): number {
  return Math.abs(row.taxAmount || row.igst + row.cgst + row.sgst);
}

export function createWorkingSet(
  pr: NormalizedInvoiceRow[],
  portal: NormalizedInvoiceRow[],
): MatchWorkingSet {
  return {
    unmatchedPr: [...pr],
    unmatchedPortal: [...portal],
    results: [],
    exceptions: [],
  };
}

/** Pass 1: exact GSTIN + normalized invoice + amount ±0.01 (or IRN). */
export function runPassExact(ws: MatchWorkingSet): void {
  const portalByKey = new Map<string, number[]>();
  for (let i = 0; i < ws.unmatchedPortal.length; i++) {
    const p = ws.unmatchedPortal[i];
    const keys = exactKeys(p);
    for (const k of keys) {
      const list = portalByKey.get(k) ?? [];
      list.push(i);
      portalByKey.set(k, list);
    }
  }

  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const keys = exactKeys(pr);
    let matchedIdx = -1;
    for (const k of keys) {
      const candidates = portalByKey.get(k) ?? [];
      for (const idx of candidates) {
        if (usedPortal.has(idx)) continue;
        const portal = ws.unmatchedPortal[idx];
        if (
          amountWithinTolerance(pr.taxableValue, portal.taxableValue, 0.01, 0) ||
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

function exactKeys(row: NormalizedInvoiceRow): string[] {
  const keys: string[] = [];
  if (row.gstin && row.normalizedInvoiceNumber) {
    keys.push(`${row.gstin}|${row.normalizedInvoiceNumber}`);
  }
  if (row.irn) {
    keys.push(`IRN|${row.irn.toUpperCase()}`);
  }
  return keys;
}

/** Pass 2: same GSTIN, fuzzy invoice, amount/date tolerances. */
export function runPassFuzzy(ws: MatchWorkingSet, settings: GstMatchSettings): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    if (!pr.gstin) {
      stillPr.push(pr);
      continue;
    }
    let bestIdx = -1;
    let bestScore = 0;
    let bestDiff = '';

    for (let i = 0; i < ws.unmatchedPortal.length; i++) {
      if (usedPortal.has(i)) continue;
      const portal = ws.unmatchedPortal[i];
      if (portal.gstin !== pr.gstin) continue;

      const sim = stringSimilarityPercent(
        pr.normalizedInvoiceNumber,
        portal.normalizedInvoiceNumber,
      );
      if (sim < settings.invoiceFuzzyThreshold) continue;
      if (
        !amountWithinTolerance(
          pr.taxableValue,
          portal.taxableValue,
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
    const amtDiff = Math.abs(pr.taxableValue - portal.taxableValue);
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
      pass: 2,
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

/** Pass 3: credit/debit notes by GSTIN + absolute amount proximity + 90 day window. */
export function runPassCdn(ws: MatchWorkingSet): void {
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const isCdn =
      pr.documentType === 'credit_note' ||
      pr.documentType === 'debit_note' ||
      pr.taxableValue < 0 ||
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
        portal.taxableValue < 0 ||
        /credit|cdn|debit/i.test(portal.narration);
      if (!portalCdn && portal.documentType !== 'amended') continue;

      const gap = Math.abs(Math.abs(pr.taxableValue) - Math.abs(portal.taxableValue));
      const dayDiff = daysBetween(pr.invoiceDate, portal.invoiceDate);
      if (dayDiff !== null && dayDiff > 90) continue;
      if (gap < bestGap) {
        bestGap = gap;
        bestIdx = i;
      }
    }

    if (bestIdx < 0 || bestGap > Math.max(1, Math.abs(pr.taxableValue) * 0.01)) {
      stillPr.push(pr);
      continue;
    }

    usedPortal.add(bestIdx);
    const portal = ws.unmatchedPortal[bestIdx];
    ws.results.push({
      status: 'CREDIT_NOTE',
      pass: 3,
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

/** Pass 4: RCM detection on remaining PR rows. */
export function runPassRcm(ws: MatchWorkingSet, settings: GstMatchSettings): void {
  if (!settings.detectRcm) return;
  const stillPr: NormalizedInvoiceRow[] = [];

  for (const pr of ws.unmatchedPr) {
    const blankGstin = !pr.gstin || !isValidGstinFormat(pr.gstin);
    const keywordHit = RCM_KEYWORDS.test(pr.narration);
    if (blankGstin || keywordHit) {
      ws.results.push({
        status: 'RCM',
        pass: 4,
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
 * Pass 5: IMS action cross-check against remaining + matched results.
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
      result.pass = result.pass ?? 5;
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
      result.pass = 5;
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
        result.pass = 5;
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
        pass: 5,
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
        pass: 5,
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
      pass: 5,
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

/** Finalize unmatched PR / portal rows. */
export function finalizeUnmatched(ws: MatchWorkingSet): void {
  for (const pr of ws.unmatchedPr) {
    ws.results.push({
      status: 'PR_ONLY',
      pass: null,
      confidence: 1,
      difference: 'In purchase register only — not found on portal',
      diffType: 'PR_ONLY',
      itcAmount: itcOf(pr),
      rcmFlag: false,
      registerRow: pr,
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
