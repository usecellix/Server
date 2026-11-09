/**
 * P3.7 / A14.3 — IRN-first matching (sales register vs GSTR-1, also safe for purchases).
 *
 * An IRN identifies one e-invoice uniquely, so when both sides carry the same IRN they are
 * the same document even if the invoice NUMBER was keyed differently. Runs before the exact
 * pass: equal amounts → MATCHED; different amounts → PARTIAL with the difference stated
 * (never silently matched). Rows without an IRN are untouched, so existing behaviour holds.
 */
import { NormalizedInvoiceRow } from '../types/domain-tool.types';
import { GstResultRow, MatchWorkingSet } from './match-passes';

const irnKey = (row: NormalizedInvoiceRow): string | null => {
  const v = String(row.irn ?? '').trim().toUpperCase();
  return v ? v : null;
};
const close = (a: number | null | undefined, b: number | null | undefined) => Math.abs((a ?? 0) - (b ?? 0)) <= 0.01;
const itc = (row: NormalizedInvoiceRow) => Math.abs(row.taxAmount || row.igst + row.cgst + row.sgst);

export function runPassIrn(ws: MatchWorkingSet): void {
  const portalByIrn = new Map<string, number>();
  ws.unmatchedPortal.forEach((p, i) => {
    const k = irnKey(p);
    if (k && !portalByIrn.has(k)) portalByIrn.set(k, i);
  });
  if (!portalByIrn.size) return;
  const usedPortal = new Set<number>();
  const stillPr: NormalizedInvoiceRow[] = [];
  for (const pr of ws.unmatchedPr) {
    const k = irnKey(pr);
    const idx = k ? portalByIrn.get(k) : undefined;
    if (idx === undefined || usedPortal.has(idx)) {
      stillPr.push(pr);
      continue;
    }
    usedPortal.add(idx);
    const portal = ws.unmatchedPortal[idx];
    const sameAmount = close(pr.taxableValue, portal.taxableValue) && close(pr.taxAmount, portal.taxAmount);
    const sameNumber = pr.normalizedInvoiceNumber === portal.normalizedInvoiceNumber;
    const notes: string[] = [];
    if (!sameNumber) notes.push(`invoice no. differs (books ${pr.invoiceNumber} / portal ${portal.invoiceNumber})`);
    if (!sameAmount) notes.push(`taxable ${pr.taxableValue ?? 0} vs ${portal.taxableValue ?? 0}, tax ${pr.taxAmount} vs ${portal.taxAmount}`);
    const result: GstResultRow = {
      status: sameAmount ? 'MATCHED' : 'PARTIAL',
      pass: 0,
      confidence: 1,
      itcAmount: itc(pr),
      rcmFlag: false,
      registerRow: pr,
      portalRow: portal,
      imsStatus: portal.imsAction ?? null,
      ...(notes.length ? { difference: `Matched on IRN; ${notes.join('; ')}`, diffType: sameAmount ? 'IRN_INVOICE_NO' : 'IRN_AMOUNT' } : {}),
    };
    ws.results.push(result);
  }
  ws.unmatchedPr = stillPr;
  ws.unmatchedPortal = ws.unmatchedPortal.filter((_, i) => !usedPortal.has(i));
}

/** GSTR-1A correction list: only the rows that need a correction (never the matched ones). */
export function gstr1aCorrections(results: GstResultRow[]): Array<{ action: string; row: GstResultRow }> {
  const out: Array<{ action: string; row: GstResultRow }> = [];
  for (const r of results) {
    if (r.status === 'MATCHED') continue;
    if (r.status === 'PARTIAL') out.push({ action: 'Amend in GSTR-1A (values differ from books)', row: r });
    else if (r.status === 'PR_ONLY') out.push({ action: 'Report missing invoice in GSTR-1A', row: r });
    else if (r.status === 'PORTAL_ONLY') out.push({ action: 'Check: reported in GSTR-1 but not in books', row: r });
  }
  return out;
}
