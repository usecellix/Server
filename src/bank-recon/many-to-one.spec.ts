/**
 * P3.9 / A18.1 — many-to-one: several ledger receipts settled by ONE bank credit
 * (the plan's test: 1,000 + 2,000 + 3,000 receipts vs a 6,000 credit). Verified against
 * the existing group pass (pass 5) in domain-tools/bank/one-to-many.ts.
 */
import { bankMatch } from '../domain-tools/bank/match-passes';
import { BankTxn } from './types';

let n = 0;
const tx = (source: 'BOOKS' | 'BANK', date: string, amount: number, direction: 'IN' | 'OUT', narration = ''): BankTxn => ({
  id: `${source}${++n}`,
  source,
  sourceFormat: 'XLSX',
  txnDate: date,
  amount,
  direction,
  narration,
  parseConfidence: 1,
  sourceLocation: { sheet: source, row: n },
});

describe('P3.9 bank recon many-to-one', () => {
  it('three ledger receipts match one bank credit as a single group', () => {
    const books = [
      tx('BOOKS', '2026-09-10', 1000, 'IN', 'Receipt Sharma Enterprises'),
      tx('BOOKS', '2026-09-10', 2000, 'IN', 'Receipt Sharma Enterprises'),
      tx('BOOKS', '2026-09-11', 3000, 'IN', 'Receipt Sharma Enterprises'),
    ];
    const bank = [tx('BANK', '2026-09-12', 6000, 'IN', 'NEFT CR SHARMA ENTERPRISES')];
    const r = bankMatch({ books, bank, periodEnd: '2026-09-30' });
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0]).toMatchObject({ kind: 'GROUP' });
    expect(r.matches[0].books.map((b) => b.amount).sort((a, b) => a - b)).toEqual([1000, 2000, 3000]);
    expect(r.unmatchedBooks).toHaveLength(0);
    expect(r.unmatchedBank).toHaveLength(0);
  });

  it('one ledger payment against several bank debits (the reverse) also groups', () => {
    const books = [tx('BOOKS', '2026-09-10', 6000, 'OUT', 'Payment Kumar Agency')];
    const bank = [
      tx('BANK', '2026-09-11', 2500, 'OUT', 'NEFT KUMAR AGENCY'),
      tx('BANK', '2026-09-12', 3500, 'OUT', 'NEFT KUMAR AGENCY'),
    ];
    const r = bankMatch({ books, bank, periodEnd: '2026-09-30' });
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0].bank).toHaveLength(2);
  });
});
