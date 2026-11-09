/**
 * P3.8 — "Compare GSTR-3B with GSTR-2B table-wise" on the active sheet.
 * Expected layout (header row): Table · 3B IGST · 3B CGST · 3B SGST · 2B IGST · 2B CGST · 2B SGST
 * (any order, any case). Output: a new report sheet with the variance per table and the tie-out.
 */
import { gstr3bVs2b, Gstr3bTableLine } from '../../domain-tools/gst/gstr3b-vs-2b.tool';
import { map3bVs2bToActions } from '../../gst-recon/gst-recon-to-actions.mapper';
import { DeterministicLane } from './lane.types';
import { indianNumber, str, toNumber } from './lane-helpers';

const findCol = (h: unknown[], ret: '3b' | '2b', tax: 'igst' | 'cgst' | 'sgst') =>
  h.findIndex((x) => {
    const t = str(x).toLowerCase().replace(/[^a-z0-9]/g, '');
    return t.includes(ret) && t.includes(tax);
  });

export function parseTable4Sheet(values: unknown[][]): Gstr3bTableLine[] | null {
  for (let r = 0; r < Math.min(values.length, 10); r += 1) {
    const h = values[r] ?? [];
    const tableCol = h.findIndex((x) => /^\s*(table|particulars|line)\b/i.test(str(x)));
    const cols = (['3b', '2b'] as const).flatMap((ret) => (['igst', 'cgst', 'sgst'] as const).map((tax) => findCol(h, ret, tax)));
    if (tableCol < 0 || cols.some((c) => c < 0)) continue;
    const lines: Gstr3bTableLine[] = [];
    for (let i = r + 1; i < values.length; i += 1) {
      const label = str(values[i]?.[tableCol]).trim();
      if (!label || /^total/i.test(label)) continue;
      const n = (c: number) => toNumber(values[i]?.[c]) ?? 0;
      lines.push({
        table: label,
        gstr3b: { igst: n(cols[0]), cgst: n(cols[1]), sgst: n(cols[2]) },
        gstr2b: { igst: n(cols[3]), cgst: n(cols[4]), sgst: n(cols[5]) },
      });
    }
    return lines.length ? lines : null;
  }
  return null;
}

export const gstr3bTableWiseLane: DeterministicLane = {
  step: 'P3.8',
  id: 'gstr3b-vs-2b-table-wise',
  writes: true,
  detect: (m) => /\b3b\b/i.test(m) && /\b2b\b/i.test(m) && /\b(table[-\s]?wise|tablewise|per table|compare)\b/i.test(m),
  run: (ctx) => {
    const tables = parseTable4Sheet(ctx.values);
    if (!tables) {
      return {
        kind: 'answer',
        answer:
          'Put the GSTR-3B Table 4 lines on this sheet with the columns **Table · 3B IGST · 3B CGST · 3B SGST · 2B IGST · 2B CGST · 2B SGST** (e.g. 4A(5) All other ITC, 4A(3) Inward RCM, 4B(1) Reversals), then ask again.',
      };
    }
    const result = gstr3bVs2b({ gstr3bItc: { igst: 0, cgst: 0, sgst: 0 }, gstr2bItc: { igst: 0, cgst: 0, sgst: 0 }, tables }).data;
    const runAt = new Date(ctx.now).toISOString().slice(0, 10);
    const sheetName = `3B vs 2B ${runAt}`.slice(0, 31);
    const lines = [
      `**GSTR-3B vs 2B, table-wise** (${tables.length} lines from ${ctx.sheetName}):`,
      ...result.tableVariances.map((t) => `- ${t.table}: ${indianNumber(t.total, t.total % 1 ? 2 : 0)}`),
      `**Tie-out:** table variances ${indianNumber(result.tieOut!.tablesTotal)} ${result.tieOut!.ties ? '=' : '≠'} total variance ${indianNumber(result.variances.total)}.`,
      `Status: ${result.status}.${result.drc01cRisk ? ` ⚠ DRC-01C flag (${result.drc01cRule}).` : ''}`,
      'Reversal lines (4B) reduce net ITC on both sides. Flags only — the decision rests with you as the professional.',
    ];
    return {
      kind: 'actions',
      answer: lines.join('\n'),
      explanation: 'Write the table-wise 3B vs 2B report to a new sheet.',
      actions: map3bVs2bToActions({ sheetName, result, runAt, relativeTo: ctx.sheetName }),
    };
  },
};

export const RECON_LANES: DeterministicLane[] = [gstr3bTableWiseLane];
