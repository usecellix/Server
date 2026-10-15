import { BankTxn, Direction } from '../../../bank-recon/types';
import { cleanText, normHeader, parseBankAmount, parseBankDate } from './bank-utils';
import { splitInstrumentAndRef } from './table-parser';

/**
 * Reads the previous period's BRS carry-forward list (items that were uncleared at the last
 * period end) and turns each into an opening books entry. They are matched against this
 * period's statement like any ledger row, so a prior-month cheque clearing in this month is
 * no longer a false "bank-only" item, and one that never clears (e.g. a stale cheque) stays
 * visible even though it never appears in this period's ledger.
 *
 * Expected block (any rows above/below are ignored): a header row containing "Type" and
 * "Amount", with optional Date, Cheque No, Party and Remarks columns.
 */
export function parseCarryForward(grid: unknown[][], sheetName: string): BankTxn[] {
  let headerIdx = -1;
  let cols: Record<string, number> = {};
  for (let i = 0; i < Math.min(grid.length, 60); i++) {
    const found: Record<string, number> = {};
    (grid[i] ?? []).forEach((c, idx) => {
      const h = normHeader(c);
      if (h === 'type') found.type = idx;
      else if (h === 'amount') found.amount = idx;
      else if (h === 'date') found.date = idx;
      else if (/^(cheque|chq|instrument)/.test(h)) found.instrument = idx;
      else if (/^(party|name)/.test(h)) found.party = idx;
      else if (/^(remarks?|narration)/.test(h)) found.remarks = idx;
    });
    if (found.type !== undefined && found.amount !== undefined && found.date !== undefined) {
      headerIdx = i;
      cols = found;
      break;
    }
  }
  if (headerIdx < 0) return [];

  const out: BankTxn[] = [];
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const row = grid[i] ?? [];
    const type = cleanText(row[cols.type]);
    const amount = parseBankAmount(row[cols.amount]);
    const date = parseBankDate(row[cols.date]);
    if (!type || !amount.present || amount.value <= 0 || !date) continue;
    let direction: Direction | null = null;
    if (/deposit|receipt|credit/i.test(type)) direction = 'IN';
    else if (/cheque|chq|payment|issued|debit/i.test(type)) direction = 'OUT';
    if (!direction) continue;

    const party = cleanText(cols.party !== undefined ? row[cols.party] : '');
    const remarks = cleanText(cols.remarks !== undefined ? row[cols.remarks] : '');
    const { instrumentNo } = splitInstrumentAndRef(cleanText(cols.instrument !== undefined ? row[cols.instrument] : ''), '');
    out.push({
      id: `BOOKS:${sheetName}:${i + 1}:CF`,
      source: 'BOOKS',
      sourceFormat: 'XLSX',
      txnDate: date,
      amount: amount.value,
      direction,
      instrumentNo,
      narration: [party, remarks].filter(Boolean).join(' — '),
      counterparty: party || undefined,
      channel: 'CHQ',
      voucherType: 'Carry-forward',
      parseConfidence: 1,
      sourceLocation: { sheet: sheetName, row: i + 1 },
    });
  }
  return out;
}
