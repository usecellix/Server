/**
 * Pricing-page features built on 2026-11-07 (from MARKETING_CLAIMS_AUDIT.md):
 *   - TDS section / rate checker  ("TDS Compliance Checks (section-wise)")
 *   - e-Invoice (IRN) validation   ("e-Invoice Validation")
 * Both read their rules from Server/src/rules/rule-tables.ts (CA to verify), add
 * check columns through the normal preview → Accept → revert pipeline, and flag —
 * never decide. Nothing in the source columns is changed.
 */
import { BatchSetOperation, SheetActionPayload } from '../types/sheet-actions.types';
import { GSTIN_STATUS_LABEL, PAN_REGEX, validateGstin } from '../utils/india-validators.util';
import { getRuleTable, ruleCitation, TdsSectionRule } from '../../rules/rule-tables';
import { DeterministicLane, LaneContext } from './lane.types';
import { columnLetter, findHeader, indianNumber, lastUsedColumn, lastUsedRow, parseDateValue, str, toNumber } from './lane-helpers';

const addr = (col: number, row: number) => `${columnLetter(col)}${row + 1}`;
const DISCLAIMER = '_Flags only, from encoded rules — the decision rests with you as the professional._';

function outputColumns(ctx: LaneContext, names: string[]): number[] {
  const h = ctx.values[ctx.headerRowIndex] ?? [];
  let next = lastUsedColumn(ctx.values) + 1;
  return names.map((n) => {
    const existing = h.findIndex((x) => str(x).trim() === n);
    return existing >= 0 ? existing : next++;
  });
}

function dataRows(ctx: LaneContext): number[] {
  const out: number[] = [];
  for (let r = ctx.headerRowIndex + 1; r <= lastUsedRow(ctx.values); r += 1) {
    if ((ctx.values[r] ?? []).some((v) => str(v).trim() !== '')) out.push(r);
  }
  return out;
}

// ---------------------------------------------------------------- TDS section / rate checker

export const normalizeSection = (raw: unknown): string =>
  str(raw).toUpperCase().replace(/^SEC(TION)?\.?\s*/i, '').replace(/\s+/g, '').trim();

export function findTdsRule(section: string, rules: TdsSectionRule[]): TdsSectionRule | undefined {
  const s = normalizeSection(section);
  return rules.find((r) => r.section.toUpperCase().replace(/\s+/g, '') === s);
}

export interface TdsRowCheck {
  expectedRate: number | null;
  expectedTds: number | null;
  issues: string[];
}

export function checkTdsRow(
  row: { section: unknown; pan: unknown; amount: unknown; rate?: unknown; tds?: unknown },
  aggregateForPartySection: number,
): TdsRowCheck {
  const table = getRuleTable('tds-sections').data;
  const issues: string[] = [];
  const rule = findTdsRule(str(row.section), table.sections);
  const amount = toNumber(row.amount);
  if (!rule) return { expectedRate: null, expectedTds: null, issues: [`Unknown section "${str(row.section) || 'blank'}"`] };
  if (amount === null) return { expectedRate: null, expectedTds: null, issues: ['Amount missing or not a number'] };

  const pan = str(row.pan).trim().toUpperCase();
  const panOk = PAN_REGEX.test(pan);
  const individual = panOk && (pan[3] === 'P' || pan[3] === 'H');
  let expectedRate = individual ? rule.rateIndividual : rule.rateOthers;
  if (!panOk) {
    expectedRate = Math.max(expectedRate, table.noPanRate);
    issues.push(`PAN ${pan ? 'invalid' : 'missing'} — higher rate ${table.noPanRate}% (s.206AA)`);
  }
  const belowSingle = rule.singleThreshold !== undefined && amount <= rule.singleThreshold;
  const belowAggregate = rule.aggregateThreshold !== undefined && aggregateForPartySection <= rule.aggregateThreshold;
  const belowThreshold =
    rule.singleThreshold !== undefined && rule.aggregateThreshold !== undefined
      ? belowSingle && belowAggregate
      : rule.singleThreshold !== undefined
        ? belowSingle
        : belowAggregate;
  const expectedTds = belowThreshold ? 0 : Math.round((amount * expectedRate) / 100);

  const statedRate = row.rate === undefined ? null : toNumber(str(row.rate).replace('%', ''));
  if (statedRate !== null && !belowThreshold) {
    const pct = statedRate > 0 && statedRate < 1 ? statedRate * 100 : statedRate; // 0.02 or 2
    if (Math.abs(pct - expectedRate) > 0.001) issues.push(`Rate ${pct}% vs ${expectedRate}% for ${rule.section} (${individual ? 'individual/HUF' : 'others'})`);
  }
  const deducted = row.tds === undefined ? null : toNumber(row.tds);
  if (deducted !== null) {
    if (belowThreshold && deducted > 0) issues.push(`Below the ${rule.section} threshold — TDS may not be needed (₹${indianNumber(deducted)} deducted)`);
    else if (!belowThreshold && deducted < expectedTds - 1) issues.push(`Short deducted by ₹${indianNumber(expectedTds - deducted)}`);
    else if (!belowThreshold && deducted > expectedTds + 1) issues.push(`Excess deducted by ₹${indianNumber(deducted - expectedTds)}`);
  }
  return { expectedRate: belowThreshold ? 0 : expectedRate, expectedTds, issues };
}

export const tdsCheckLane: DeterministicLane = {
  step: 'MKT-TDS',
  id: 'tds-check',
  writes: true,
  detect: (m) => /\btds\b/i.test(m) && /\b(check|verify|validate|review|audit|compliance|correct|rates?|sections?)\b/i.test(m) && !/\b26\s*as\b|\bdropdown\b|\btemplate\b|\bworking\b/i.test(m),
  run: (ctx) => {
    const h = ctx.values[ctx.headerRowIndex] ?? [];
    const cSection = findHeader(h, /\bsection\b|\bsec\b/i);
    const cAmount = findHeader(h, /amount\s*paid|payment|gross|amount|value/i);
    if (cSection < 0 || cAmount < 0) {
      return {
        kind: 'answer',
        answer: 'I need at least a **Section** column and an **Amount** column (plus PAN, TDS Rate and TDS Deducted if you have them). Example: Deductee · PAN · Section · Amount Paid · TDS Rate · TDS Deducted.',
      };
    }
    const cPan = findHeader(h, /\bpan\b/i);
    const cParty = findHeader(h, /deductee|party|vendor|name|payee/i);
    const cRate = findHeader(h, /rate/i);
    const cTds = findHeader(h, /tds\s*(deducted|amount)?$|deducted/i);
    const rows = dataRows(ctx);

    // Aggregate per deductee (PAN, else name) + section across the sheet, for aggregate thresholds.
    const aggregate = new Map<string, number>();
    const keyOf = (r: number) => `${str(ctx.values[r]?.[cPan >= 0 ? cPan : cParty]).trim().toUpperCase()}|${normalizeSection(ctx.values[r]?.[cSection])}`;
    for (const r of rows) aggregate.set(keyOf(r), (aggregate.get(keyOf(r)) ?? 0) + (toNumber(ctx.values[r]?.[cAmount]) ?? 0));

    const [cExpRate, cExpTds, cCheck] = outputColumns(ctx, ['Expected Rate %', 'Expected TDS', 'TDS Check']);
    const ops: BatchSetOperation[] = [
      { address: addr(cExpRate, ctx.headerRowIndex), value: 'Expected Rate %' },
      { address: addr(cExpTds, ctx.headerRowIndex), value: 'Expected TDS' },
      { address: addr(cCheck, ctx.headerRowIndex), value: 'TDS Check' },
    ];
    const flagged: string[] = [];
    for (const r of rows) {
      const res = checkTdsRow(
        {
          section: ctx.values[r]?.[cSection],
          pan: cPan >= 0 ? ctx.values[r]?.[cPan] : '',
          amount: ctx.values[r]?.[cAmount],
          rate: cRate >= 0 && cRate !== cExpRate ? ctx.values[r]?.[cRate] : undefined,
          tds: cTds >= 0 && cTds !== cExpTds && cTds !== cCheck ? ctx.values[r]?.[cTds] : undefined,
        },
        aggregate.get(keyOf(r)) ?? 0,
      );
      ops.push({ address: addr(cExpRate, r), value: res.expectedRate ?? '' });
      ops.push({ address: addr(cExpTds, r), value: res.expectedTds ?? '' });
      ops.push({ address: addr(cCheck, r), value: res.issues.length ? res.issues.join('; ') : 'OK' });
      if (res.issues.length) flagged.push(`row ${r + 1}${cParty >= 0 ? ` (${str(ctx.values[r]?.[cParty])})` : ''} — ${res.issues.join('; ')}`);
    }
    return {
      kind: 'actions',
      answer: [
        `Checked **${rows.length}** TDS rows: **${rows.length - flagged.length} OK**, **${flagged.length} flagged**. I'll add Expected Rate %, Expected TDS and TDS Check; your columns are not changed.`,
        ...(flagged.length ? ['', ...flagged.slice(0, 15).map((f) => `- ${f}`)] : []),
        ...(flagged.length > 15 ? [`- …and ${flagged.length - 15} more`] : []),
        '',
        `**Rule source:** ${ruleCitation('tds-sections')}.`,
        DISCLAIMER,
      ].join('\n'),
      explanation: `Check TDS sections and rates for ${rows.length} rows.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- e-Invoice (IRN) validation

export function checkEInvoiceRow(
  row: { irn?: unknown; ackNo?: unknown; docNo?: unknown; docDate?: unknown; sellerGstin?: unknown; buyerGstin?: unknown; taxable?: unknown },
  today: Date,
): string[] {
  const rule = getRuleTable('e-invoice').data;
  const issues: string[] = [];
  if (row.irn !== undefined) {
    const irn = str(row.irn).trim();
    if (!irn) issues.push('IRN missing');
    else if (!new RegExp(`^[0-9a-fA-F]{${rule.irnLength}}$`).test(irn)) issues.push(`IRN must be ${rule.irnLength} hex characters (has ${irn.length})`);
  }
  if (row.ackNo !== undefined && str(row.ackNo).trim() !== '') {
    const ack = str(row.ackNo).trim().replace(/\.0+$/, '');
    if (!new RegExp(`^\\d{${rule.ackNoLength}}$`).test(ack)) issues.push(`Ack No. must be ${rule.ackNoLength} digits`);
  }
  if (row.docNo !== undefined) {
    const doc = str(row.docNo).trim();
    if (!doc) issues.push('Invoice no. missing');
    else if (doc.length > rule.docNoMaxLength) issues.push(`Invoice no. longer than ${rule.docNoMaxLength} characters`);
    else if (!new RegExp(rule.docNoPattern).test(doc)) issues.push('Invoice no. has a character not allowed (only letters, digits, / and -; cannot start with 0, / or -)');
  }
  if (row.docDate !== undefined) {
    const ms = parseDateValue(row.docDate);
    const todayMs = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
    if (ms === null) issues.push('Invoice date missing or not a date');
    else if (ms > todayMs) issues.push('Invoice date is in the future');
  }
  const seller = str(row.sellerGstin).trim();
  const buyer = str(row.buyerGstin).trim();
  if (row.sellerGstin !== undefined) {
    if (!seller) issues.push('Supplier GSTIN missing');
    else if (validateGstin(seller).status !== 'valid') issues.push(`Supplier GSTIN: ${GSTIN_STATUS_LABEL[validateGstin(seller).status]}`);
  }
  if (row.buyerGstin !== undefined && buyer && buyer.toUpperCase() !== 'URP') {
    if (validateGstin(buyer).status !== 'valid') issues.push(`Buyer GSTIN: ${GSTIN_STATUS_LABEL[validateGstin(buyer).status]}`);
    else if (seller && buyer.toUpperCase() === seller.toUpperCase()) issues.push('Buyer GSTIN equals supplier GSTIN');
  }
  if (row.taxable !== undefined) {
    const t = toNumber(row.taxable);
    if (t === null) issues.push('Taxable value missing or not a number');
    else if (t < 0) issues.push('Taxable value is negative (use a credit note)');
  }
  return issues;
}

export const eInvoiceLane: DeterministicLane = {
  step: 'MKT-EINV',
  id: 'e-invoice',
  writes: true,
  detect: (m) =>
    /\b(e-?invoices?|irns?|einvoic\w*)\b/i.test(m) && /\b(validat\w*|check|verify|review|audit)\b/i.test(m) && !/\breconcil\w*\b/i.test(m),
  run: (ctx) => {
    const h = ctx.values[ctx.headerRowIndex] ?? [];
    const col = (re: RegExp) => findHeader(h, re);
    const c = {
      irn: col(/^\s*irn\b|invoice reference/i),
      ack: col(/ack(nowledg\w*)?\.?\s*no/i),
      doc: col(/(invoice|doc(ument)?|bill)\s*(no|number|#)/i),
      date: col(/(invoice|doc(ument)?|bill)?\s*date/i),
      seller: col(/(supplier|seller|from)\s*gstin|^gstin$/i),
      buyer: col(/(buyer|recipient|customer|to)\s*gstin/i),
      taxable: col(/taxable/i),
    };
    if (c.irn < 0 && c.doc < 0) {
      return { kind: 'answer', answer: 'I need an **IRN** or **Invoice No** column to validate e-invoices (plus Invoice Date, Supplier GSTIN, Buyer GSTIN, Taxable Value if you have them).' };
    }
    const rows = dataRows(ctx);
    const get = (r: number, i: number) => (i >= 0 ? ctx.values[r]?.[i] : undefined);
    const [cStatus, cIssues] = outputColumns(ctx, ['e-Invoice Check', 'e-Invoice Issues']);
    const ops: BatchSetOperation[] = [
      { address: addr(cStatus, ctx.headerRowIndex), value: 'e-Invoice Check' },
      { address: addr(cIssues, ctx.headerRowIndex), value: 'e-Invoice Issues' },
    ];
    const seenIrn = new Map<string, number>();
    const seenDoc = new Map<string, number>();
    const flagged: string[] = [];
    for (const r of rows) {
      const issues = checkEInvoiceRow(
        { irn: get(r, c.irn), ackNo: get(r, c.ack), docNo: get(r, c.doc), docDate: get(r, c.date), sellerGstin: get(r, c.seller), buyerGstin: get(r, c.buyer), taxable: get(r, c.taxable) },
        ctx.now,
      );
      const irn = str(get(r, c.irn)).trim().toLowerCase();
      if (irn) {
        if (seenIrn.has(irn)) issues.push(`Duplicate IRN (same as row ${seenIrn.get(irn)! + 1})`);
        else seenIrn.set(irn, r);
      }
      const docKey = `${str(get(r, c.seller)).trim().toUpperCase()}|${str(get(r, c.doc)).trim().toUpperCase()}`;
      if (c.doc >= 0 && str(get(r, c.doc)).trim()) {
        if (seenDoc.has(docKey)) issues.push(`Duplicate invoice no. for this supplier (row ${seenDoc.get(docKey)! + 1})`);
        else seenDoc.set(docKey, r);
      }
      ops.push({ address: addr(cStatus, r), value: issues.length ? 'Check' : 'OK' });
      ops.push({ address: addr(cIssues, r), value: issues.join('; ') });
      if (issues.length) flagged.push(`row ${r + 1} — ${issues.join('; ')}`);
    }
    return {
      kind: 'actions',
      answer: [
        `Validated **${rows.length}** e-invoices: **${rows.length - flagged.length} OK**, **${flagged.length} to check**. I'll add e-Invoice Check and e-Invoice Issues; nothing else changes.`,
        ...(flagged.length ? ['', ...flagged.slice(0, 15).map((f) => `- ${f}`)] : []),
        ...(flagged.length > 15 ? [`- …and ${flagged.length - 15} more`] : []),
        '',
        `**Rule source:** ${ruleCitation('e-invoice')}.`,
        '_Format checks only — CELLIX cannot confirm an IRN is registered on the IRP portal._',
      ].join('\n'),
      explanation: `Validate ${rows.length} e-invoices.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

export const COMPLIANCE_LANES: DeterministicLane[] = [tdsCheckLane, eInvoiceLane];
