import { DomainTool, DomainToolResult } from '../types/domain-tool.types';
import { roundMoney } from './normalize-invoice';

export interface Gstr3bComponent {
  igst: number;
  cgst: number;
  sgst: number;
}

export interface Gstr3bVs2bInput {
  /** ITC claimed in GSTR-3B Table 4 (net or available — as provided by CA) */
  gstr3bItc: Gstr3bComponent;
  /** ITC available from GSTR-2B totals */
  gstr2bItc: Gstr3bComponent;
  /** DRC-01C threshold: excess of lower of ₹5L or 10% — defaults applied */
  excessAbsThreshold?: number;
  excessPctThreshold?: number;
}

export interface Gstr3bVs2bOutput {
  variances: {
    igst: number;
    cgst: number;
    sgst: number;
    total: number;
  };
  excessClaim: number;
  shortClaim: number;
  drc01cRisk: boolean;
  status: 'BALANCED' | 'EXCESS' | 'SHORT' | 'MIXED';
  messages: string[];
}

/**
 * Summary-level GSTR-3B vs GSTR-2B ITC comparison (not invoice matching).
 */
export const gstr3bVs2b: DomainTool<Gstr3bVs2bInput, Gstr3bVs2bOutput> = (input) => {
  const a = input.gstr3bItc;
  const b = input.gstr2bItc;
  const igst = roundMoney(a.igst - b.igst);
  const cgst = roundMoney(a.cgst - b.cgst);
  const sgst = roundMoney(a.sgst - b.sgst);
  const total = roundMoney(igst + cgst + sgst);

  const excessClaim = Math.max(0, total);
  const shortClaim = Math.max(0, -total);

  const absTh = input.excessAbsThreshold ?? 500000;
  const pctTh = input.excessPctThreshold ?? 10;
  const availableTotal = Math.abs(b.igst + b.cgst + b.sgst) || 1;
  const excessPct = (excessClaim / availableTotal) * 100;
  const drc01cRisk =
    excessClaim > 0 && (excessClaim >= absTh || excessPct >= pctTh);

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
  if (drc01cRisk) {
    messages.push(
      'DRC-01C risk: excess exceeds ₹5L or 10% of available ITC (whichever lower) threshold proxy.',
    );
  }

  const result: DomainToolResult<Gstr3bVs2bOutput> = {
    data: {
      variances: { igst, cgst, sgst, total },
      excessClaim: roundMoney(excessClaim),
      shortClaim: roundMoney(shortClaim),
      drc01cRisk,
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
