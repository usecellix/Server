import { gstr3bVs2b } from './gstr3b-vs-2b.tool';
import { map3bVs2bToActions } from '../../gst-recon/gst-recon-to-actions.mapper';

const c = (igst: number, cgst = 0, sgst = 0) => ({ igst, cgst, sgst });

describe('P3.8 GSTR-3B vs 2B, table-wise', () => {
  const tables = [
    { table: '4A(5) All other ITC', gstr3b: c(100000, 50000, 50000), gstr2b: c(95000, 50000, 50000) },
    { table: '4A(3) Inward RCM', gstr3b: c(18000), gstr2b: c(0) },
    { table: '4B(1) Reversals (Rule 42/43)', gstr3b: c(2000), gstr2b: c(0) },
  ];
  const res = gstr3bVs2b({ gstr3bItc: c(0), gstr2bItc: c(0), tables }).data;

  it('variance per table, reversals reduce net ITC', () => {
    expect(res.tableVariances.map((t) => [t.table, t.total])).toEqual([
      ['4A(5) All other ITC', 5000],
      ['4A(3) Inward RCM', 18000],
      ['4B(1) Reversals (Rule 42/43)', -2000],
    ]);
  });

  it('sum of table variances = total variance', () => {
    expect(res.tieOut).toEqual({ tablesTotal: 21000, overallTotal: 21000, ties: true });
    expect(res.variances.total).toBe(21000);
  });

  it('report sheet lists each table and the tie-out', () => {
    const actions = map3bVs2bToActions({ sheetName: '3B', result: res, runAt: '07-11-2026' });
    const rows = (actions.find((a) => a.type === 'WRITE_TABLE') as { rows: unknown[][] }).rows;
    expect(rows.some((r) => r[0] === 'Table 4A(3) Inward RCM' && r[4] === 18000)).toBe(true);
    expect(rows.some((r) => r[0] === 'Tie-out' && /= overall 21000/.test(String(r[5])))).toBe(true);
  });
});

describe('P4.4 DRC-01C threshold from the rule table', () => {
  it('flags only when the excess is above BOTH ₹25 lakh and 20% (rule 88D), and cites the rule', () => {
    const big = gstr3bVs2b({ gstr3bItc: c(5_000_000), gstr2bItc: c(2_000_000) }).data; // 30L excess, 150%
    expect(big.drc01cRisk).toBe(true);
    expect(big.drc01cRule).toMatch(/drc-01c v2026-11\.v1 \(CA to verify\)/);
    const pctOnly = gstr3bVs2b({ gstr3bItc: c(130_000), gstr2bItc: c(100_000) }).data; // 30%, ₹30k
    expect(pctOnly.drc01cRisk).toBe(false);
    const amountOnly = gstr3bVs2b({ gstr3bItc: c(105_000_000), gstr2bItc: c(100_000_000) }).data; // ₹50L, 5%
    expect(amountOnly.drc01cRisk).toBe(false);
  });
  it('explicit thresholds still override', () => {
    expect(gstr3bVs2b({ gstr3bItc: c(130_000), gstr2bItc: c(100_000), excessAbsThreshold: 10_000, excessPctThreshold: 10 }).data.drc01cRisk).toBe(true);
  });
});
