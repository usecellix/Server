import { runBankReconEngine } from './bank-recon.engine';
// eslint-disable-next-line @typescript-eslint/no-require-imports
const fx = require('./__fixtures__/demo-sep2026.json') as Record<string, unknown[][]>;

/** Sheets of cellix_bank_recon_demo_workbook_sep2026.xlsx: sheet1 = Bank Ledger, sheet2 = Bank Statement, sheet3 = Prior BRS. */
const run = (withPrior: boolean, period = true) =>
  runBankReconEngine({
    books: { sheetName: 'Bank Ledger', data: fx.sheet1 },
    bank: { sheetName: 'Bank Statement', data: fx.sheet2 },
    priorBrs: withPrior ? { sheetName: 'Prior BRS (Aug-26)', data: fx.sheet3 } : undefined,
    ...(period ? { periodStart: '2026-09-01', periodEnd: '2026-09-30' } : {}),
  });

const count = (items: Array<{ reason: string }>) =>
  items.reduce<Record<string, number>>((r, u) => ((r[u.reason] = (r[u.reason] ?? 0) + 1), r), {});

describe('demo workbook (Sep 2026)', () => {
  it('reads every row and raises no false validation errors (Tally closing balance sits in the opposite column)', () => {
    const o = run(true);
    expect(o.booksParsed.txns).toHaveLength(57);
    expect(o.bankParsed.txns).toHaveLength(65);
    expect(o.booksValidation.errors).toEqual([]);
    expect(o.bankValidation.errors).toEqual([]);
    expect(o.booksParsed.closingBalance).toBe(2810451.75);
  });

  it('infers the stated period from the sheet headers when none is given', () => {
    const o = run(true, false);
    expect(o.periodStart).toBe('2026-09-01');
    expect(o.periodEnd).toBe('2026-09-30');
  });

  it('with the prior BRS carried forward: 43 matches + 2 carry-forward, 30 unreconciled, BRS ties to 0.00', () => {
    const o = run(true);
    expect(o.carryForwardCount).toBe(3);
    const cf = o.result.matches.filter((m) => m.books.some((b) => b.voucherType === 'Carry-forward'));
    expect(cf).toHaveLength(2);
    expect(o.result.matches.length - cf.length).toBe(43);

    const unreconciled = o.result.unmatchedBooks.length + o.result.unmatchedBank.length;
    expect(unreconciled).toBe(30);
    expect(o.result.unmatchedBooks.find((u) => u.reason === 'stale_cheque')?.txn.amount).toBe(15750);

    expect(o.brs.difference).toBe(0);
    expect(o.brs.tiedOut).toBe(true);
    expect(o.brs.booksBalance).toBe(2810451.75);
    expect(o.brs.bankBalance).toBe(2478981.05);
    expect(o.brs.adjustedBooks).toBe(2849601.05);
    expect(o.brs.adjustedBank).toBe(2849601.05);
  });

  it('without the prior BRS it does not tie out and says why (residual = the prior-month items)', () => {
    const o = run(false);
    expect(o.brs.tiedOut).toBe(false);
    expect(o.brs.difference).toBe(7750);
  });

  it('diagnoses each designed scenario', () => {
    const o = run(true);
    const books = count(o.result.unmatchedBooks);
    const bank = count(o.result.unmatchedBank);
    expect(books).toMatchObject({
      amount_mismatch: 1,
      date_mismatch_outside_window: 1,
      duplicate_in_books: 1,
      genuinely_missing: 2, // NEFT payment + NEFT receipt with a UTR but no bank trace
      stale_cheque: 1,
      cheque_issued_not_cleared: 2,
      deposit_in_transit: 2,
    });
    expect(bank).toMatchObject({
      possible_duplicate_in_bank: 1,
      return_credit_not_booked: 1,
      cheque_bounce_not_booked: 3,
      unidentified_transaction: 1,
      interest_credited_not_booked: 1,
      tds_on_interest_not_booked: 1,
    });
    expect(bank.direct_debit_not_booked).toBe(3); // EMI, insurance, GST challan
    expect(bank.direct_credit_not_booked).toBe(3);
    expect(bank.bank_charges_not_booked).toBe(6);
  });
});

describe('service: "not cleared" view', () => {
  it('lists only timing items, on its own sheet, while the full result is untouched', async () => {
    const { BankReconService } = await import('./bank-recon.service');
    const base = {
      books_ledger: { sheet_name: 'Bank Ledger', data: fx.sheet1 },
      bank_statement: { sheet_name: 'Bank Statement', data: fx.sheet2 },
      prior_brs: { sheet_name: 'Prior BRS (Aug-26)', data: fx.sheet3 },
      casual: true,
      layout: 'books_flat' as const,
    };
    const svc = new BankReconService();
    const all: any = svc.reconcile(base as any);
    const timing: any = svc.reconcile({ ...base, books_filter: 'timing' } as any);
    expect(all.rows.filter((r: any) => r.side === 'BOOKS')).toHaveLength(10);
    expect(timing.rows.filter((r: any) => r.side === 'BOOKS')).toHaveLength(5); // 2 cheques, 2 deposits in transit, 1 stale cheque
    expect(timing.output_sheet_name).toBe('Bank Recon - Not Cleared');
    expect(timing.brs.tiedOut).toBe(true);
  });
});
