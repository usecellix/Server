import { DomainTool, DomainToolResult } from '../types/domain-tool.types';
import { roundMoney } from './normalize-invoice';
import { getRuleTable, ruleCitation } from '../../rules/rule-tables';

export interface Gstr3bComponent {
  igst: number;
  cgst: number;
  sgst: number;
}

/** P3.8 — one GSTR-3B Table 4 line compared with its GSTR-2B counterpart. */
export interface Gstr3bTableLine {
  /** e.g. "4A(5) All other ITC", "4A(3) Inward RCM", "4B(1) Reversals (Rule 42/43)", "4B(2) Other reversals" */
  table: string;
  gstr3b: Gstr3bComponent;
  gstr2b: Gstr3bComponent;
}

export interface Gstr3bVs2bInput {
  /** ITC claimed in GSTR-3B Table 4 (net or available — as provided by CA) */
  gstr3bItc: Gstr3bComponent;
  /** ITC available from GSTR-2B totals */
  gstr2bItc: Gstr3bComponent;
  /** Optional table-wise lines (4A incl. RCM, 4B reversals). When given, totals are their sum. */
  tables?: Gstr3bTableLine[];
  /** Override the DRC-01C thresholds (default: rule table `drc-01c`). */
  excessAbsThreshold?: number;
  excessPctThreshold?: number;
}

export interface Gstr3bTableVariance {
  table: string;
  igst: number;
  cgst: number;
  sgst: number;
  total: number;
}

export interface Gstr3bVs2bOutput {
  variances: {
    igst: number;
    cgst: number;
    sgst: number;
    total: number;
  };
  /** P3.8 — per Table 4 line; their totals add up to `variances.total` (see `tieOut`). */
  tableVariances: Gstr3bTableVariance[];
  tieOut: { tablesTotal: number; overallTotal: number; ties: boolean } | null;
  excessClaim: number;
  shortClaim: number;
  drc01cRisk: boolean;
  /** P4.4 — which rule decided the DRC-01C flag. */
  drc01cRule: string;
  status: 'BALANCED' | 'EXCESS' | 'SHORT' | 'MIXED';
  messages: string[];
}

const sum = (c: Gstr3bComponent) => c.igst + c.cgst + c.sgst;
const add = (a: Gstr3bComponent, b: Gstr3bComponent): Gstr3bComponent => ({ igst: a.igst + b.igst, cgst: a.cgst + b.cgst, sgst: a.sgst + b.sgst });
/** Reversal lines (4B) reduce the net ITC on both sides. */
const signFor = (table: string) => (/^\s*4B/i.test(table) ? -1 : 1);

/**
 * Summary-level GSTR-3B vs GSTR-2B ITC comparison (not invoice matching), optionally table-wise.
 */
export const gstr3bVs2b: DomainTool<Gstr3bVs2bInput, Gstr3bVs2bOutput> = (input) => {
  let a = input.gstr3bItc;
  let b = input.gstr2bItc;
  const tableVariances: Gstr3bTableVariance[] = [];
  if (input.tables?.length) {
    const zero = { igst: 0, cgst: 0, sgst: 0 };
    let net3b = zero;
    let net2b = zero;
    for (const line of input.tables) {
      const s = signFor(line.table);
      const l3 = { igst: s * line.gstr3b.igst, cgst: s * line.gstr3b.cgst, sgst: s * line.gstr3b.sgst };
      const l2 = { igst: s * line.gstr2b.igst, cgst: s * line.gstr2b.cgst, sgst: s * line.gstr2b.sgst };
      net3b = add(net3b, l3);
      net2b = add(net2b, l2);
      const v = { igst: roundMoney(l3.igst - l2.igst), cgst: roundMoney(l3.cgst - l2.cgst), sgst: roundMoney(l3.sgst - l2.sgst) };
      tableVariances.push({ table: line.table, ...v, total: roundMoney(sum(v)) });
    }
    a = net3b;
    b = net2b;
  }
  const igst = roundMoney(a.igst - b.igst);
  const cgst = roundMoney(a.cgst - b.cgst);
  const sgst = roundMoney(a.sgst - b.sgst);
  const total = roundMoney(igst + cgst + sgst);
  const tablesTotal = roundMoney(tableVariances.reduce((s, t) => s + t.total, 0));
  const tieOut = tableVariances.length ? { tablesTotal, overallTotal: total, ties: Math.abs(tablesTotal - total) < 0.01 } : null;

  const excessClaim = Math.max(0, total);
  const shortClaim = Math.max(0, -total);

  // P4.4 — thresholds from the rule table (Rule 88D: excess above ₹ amount AND % of 2B ITC).
  const rule = getRuleTable('drc-01c').data;
  const absTh = input.excessAbsThreshold ?? rule.amountRupees;
  const pctTh = input.excessPctThreshold ?? rule.percent;
  const availableTotal = Math.abs(sum(b)) || 1;
  const excessPct = (excessClaim / availableTotal) * 100;
  const drc01cRisk = excessClaim > 0 && excessClaim > absTh && excessPct > pctTh;

  let status: Gstr3bVs2bOutput['status'] = 'BALANCED';
  if (Math.abs(total) < 0.01) status = 'BALANCED';
  else if (excessClaim > 0 && shortClaim === 0) status = 'EXCESS';
  else if (shortClaim > 0 && excessClaim === 0) status = 'SHORT';
  else status = 'MIXED';

  const messages: string[] = [];
  if (status === 'BALANCED') {
    messages.push('GSTR-3B ITC claims align with GSTR-2B available ITC (within ₹0.01).');
  }
  if (excessClaim > 0) {
    messages.push(`Excess ITC claimed vs 2B: ₹${excessClaim.toFixed(2)}.`);
  }
  if (shortClaim > 0) {
    messages.push(`Short ITC claimed vs 2B: ₹${shortClaim.toFixed(2)} (possible under-claim).`);
  }
  if (tieOut) {
    messages.push(`Table-wise variances total ₹${tablesTotal.toFixed(2)} ${tieOut.ties ? '=' : '≠'} overall variance ₹${total.toFixed(2)}.`);
  }
  const drc01cRule = ruleCitation('drc-01c');
  if (drc01cRisk) {
    messages.push(`DRC-01C flag: excess exceeds ₹${absTh.toLocaleString('en-IN')} and ${pctTh}% of 2B ITC (${drc01cRule}).`);
  }

  const result: DomainToolResult<Gstr3bVs2bOutput> = {
    data: {
      variances: { igst, cgst, sgst, total },
      tableVariances,
      tieOut,
      excessClaim: roundMoney(excessClaim),
      shortClaim: roundMoney(shortClaim),
      drc01cRisk,
      drc01cRule,
      status,
      messages,
    },
    confidence: 1,
    exceptions: drc01cRisk
      ? [
          {
            code: 'GST_DRC01C_RISK',
            severity: 'flag',
            message: messages[messages.length - 1] ?? 'DRC-01C risk',
            affectedRows: [],
          },
        ]
      : [],
    sourceRefs: [],
  };
  return result;
};
