import { runBankReconEngine } from './bank-recon.engine';
import { parseBankTable } from '../domain-tools/ingestion/bank/table-parser';
import { parseBankDate, parseBankAmount } from '../domain-tools/ingestion/bank/bank-utils';
import { parseNarration } from '../domain-tools/bank/narration-parser';
import { bankMatch } from '../domain-tools/bank/match-passes';
import { mapBankReconToActions } from './bank-recon-to-actions.mapper';
import { BankTxn } from './types';

const BANK_HEADER = ['Date', 'Narration', 'Chq./Ref.No.', 'Value Dt', 'Withdrawal Amt.', 'Deposit Amt.', 'Closing Balance'];
const bankRows: unknown[][] = [
  ['02/04/25', 'NEFT-HDFCN52025040211-ABC TRADERS-INV 4521', 'HDFCN52025040211', '02/04/25', 20000, '', 80000],
  ['05/04/25', 'CHQ DEP 000123', '000123', '05/04/25', '', 50000, 130000],
  ['10/04/25', 'UPI/412345678901/RAMESH/pay', '412345678901', '10/04/25', 1500, '', 128500],
  ['12/04/25', 'CHQ 000456', '000456', '12/04/25', 30000, '', 98500],
  ['30/04/25', 'INT PD 01-01-25 TO 31-03-25', '', '30/04/25', '', 850, 99350],
  ['30/04/25', 'SMS CHARGES', '', '30/04/25', 118, '', 99232],
  ['05/05/25', 'CHQ 000789', '000789', '05/05/25', 10000, '', 89232],
];
const bankPreamble: unknown[][] = [
  ['HDFC BANK LTD'],
  ['Account No : 50100123456789   IFSC: HDFC0000123'],
  [''],
];
const bankGrid = (rows = bankRows, opening = 100000): unknown[][] => [
  ...bankPreamble,
  BANK_HEADER,
  ['Opening Balance', '', '', '', '', '', opening],
  ...rows,
];

const BOOKS_HEADER = ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Debit', 'Credit'];
const booksRows: unknown[][] = [
  ['02-04-2025', 'ABC Traders', 'Payment', 1, '', 20000],
  ['04-04-2025', 'XYZ Corp Chq 123', 'Receipt', 2, 50000, ''],
  ['08-04-2025', 'Ramesh', 'Payment', 3, '', 1500],
  ['11-04-2025', 'Vendor P cheque no 456', 'Payment', 4, '', 30000],
  ['28-04-2025', 'Vendor Q cheque no 789', 'Payment', 5, '', 10000],
  ['29-04-2025', 'DEF Ltd cheque no 555', 'Receipt', 6, 7000, ''],
];
const booksGrid = (rows = booksRows): unknown[][] => [
  ['Bank Account - HDFC'],
  BOOKS_HEADER,
  ['Opening Balance', '', '', '', 100000, ''],
  ...rows,
];

describe('bank-utils', () => {
  it('parses Indian dates, serials and month names', () => {
    expect(parseBankDate('02/04/25')).toBe('2025-04-02');
    expect(parseBankDate('02-Apr-2025')).toBe('2025-04-02');
    expect(parseBankDate('2025-04-02')).toBe('2025-04-02');
    expect(parseBankDate(45749)).toBe('2025-04-02');
    expect(parseBankDate('31/02/25')).toBe('');
    expect(parseBankDate('Narration')).toBe('');
  });
  it('parses amounts with separators and Cr/Dr suffix', () => {
    expect(parseBankAmount('1,23,456.50').value).toBe(123456.5);
    expect(parseBankAmount('2,000.00 Cr')).toMatchObject({ value: 2000, suffix: 'CR' });
    expect(parseBankAmount('(500)').value).toBe(-500);
    expect(parseBankAmount('').present).toBe(false);
  });
});

describe('narration-parser', () => {
  it('extracts channel, refs and counterparty', () => {
    expect(parseNarration('NEFT-HDFCN52025040211-ABC TRADERS-INV 4521')).toMatchObject({
      channel: 'NEFT',
      referenceNo: 'HDFCN52025040211',
      counterparty: 'ABC TRADERS',
    });
    expect(parseNarration('UPI/412345678901/RAMESH/pay')).toMatchObject({ channel: 'UPI', referenceNo: '412345678901' });
    expect(parseNarration('CHQ DEP 000123')).toMatchObject({ channel: 'CHQ', instrumentNo: '123' });
  });
});

describe('table-parser', () => {
  it('skips preamble, reads opening balance, maps bank Credit → IN', () => {
    const p = parseBankTable(bankGrid(), { role: 'BANK', sheetName: 'Stmt' });
    expect(p.headerRow).toBe(4);
    expect(p.openingBalance).toBe(100000);
    expect(p.txns).toHaveLength(7);
    expect(p.txns[1]).toMatchObject({ direction: 'IN', amount: 50000, instrumentNo: '123' });
    expect(p.txns[0]).toMatchObject({ direction: 'OUT', referenceNo: 'HDFCN52025040211' });
    expect(p.closingBalance).toBe(89232);
  });

  it('maps books Debit → IN (receipt) and Credit → OUT — the sign convention', () => {
    const p = parseBankTable(booksGrid(), { role: 'BOOKS', sheetName: 'Ledger' });
    expect(p.openingBalance).toBe(100000);
    const receipt = p.txns.find((t) => t.voucherNo === '2')!;
    const payment = p.txns.find((t) => t.voucherNo === '1')!;
    expect(receipt.direction).toBe('IN');
    expect(payment.direction).toBe('OUT');
  });

  it('handles newest-first statements and a single Amount + Dr/Cr column', () => {
    const grid = [
      ['Date', 'Description', 'Amount', 'Dr/Cr', 'Balance'],
      ['03/04/25', 'Second', 500, 'Dr', 1500],
      ['02/04/25', 'First', 1000, 'Cr', 2000],
      ['01/04/25', 'Opening Balance', '', '', 1000],
    ];
    const p = parseBankTable(grid, { role: 'BANK', sheetName: 'S' });
    expect(p.txns.map((t) => t.narration)).toEqual(['First', 'Second']);
    expect(p.txns.map((t) => t.direction)).toEqual(['IN', 'OUT']);
    expect(p.openingBalance).toBe(1000);
  });

  it('appends wrapped narration lines and never drops a dated row with an amount', () => {
    const grid = [
      ['Date', 'Narration', 'Debit', 'Credit', 'Balance'],
      ['02/04/25', 'NEFT PAYMENT TO', 100, '', 900],
      ['', 'ABC TRADERS CONTINUED', '', '', ''],
      ['03/04/25', 'Total', 5, '', 895],
    ];
    const p = parseBankTable(grid, { role: 'BANK', sheetName: 'S' });
    expect(p.txns).toHaveLength(2);
    expect(p.txns[0].narration).toBe('NEFT PAYMENT TO ABC TRADERS CONTINUED');
  });
});

describe('runBankReconEngine', () => {
  const run = (bank = bankGrid(), books = booksGrid()) =>
    runBankReconEngine({
      books: { sheetName: 'Ledger', data: books },
      bank: { sheetName: 'Stmt', data: bank },
      periodStart: '2025-04-01',
      periodEnd: '2025-04-30',
    });

  it('matches by reference / amount+date and ties the BRS out to zero', () => {
    const out = run();
    expect(out.booksValidation.errors).toEqual([]);
    expect(out.bankValidation.errors).toEqual([]);

    const r = out.result;
    expect(r.matches).toHaveLength(5);
    // cheque 456 matched on reference (pass 1); NEFT matched on amount+date (pass 2)
    expect(r.matches.find((m) => m.books[0].voucherNo === '4')!.pass).toBe(1);
    expect(r.matches.find((m) => m.books[0].voucherNo === '1')!.pass).toBe(2);
    // cheque 789 clears after period end but still matches (clearing window)
    expect(r.matches.find((m) => m.books[0].voucherNo === '5')).toBeDefined();

    expect(r.unmatchedBooks.map((u) => [u.txn.voucherNo, u.reason])).toEqual([['6', 'deposit_in_transit']]);
    expect(r.unmatchedBank.map((u) => u.reason).sort()).toEqual(
      ['bank_charges_not_booked', 'interest_credited_not_booked'].sort(),
    );

    expect(out.brs.booksBalance).toBe(95500);
    expect(out.brs.bankBalance).toBe(99232);
    expect(out.brs.booksOnlyIn).toBe(7000); // deposit in transit
    expect(out.brs.booksOnlyOut).toBe(10000); // cheque 789 cleared 05-May
    expect(out.brs.bankOnlyIn).toBe(850);
    expect(out.brs.bankOnlyOut).toBe(118);
    expect(out.brs.difference).toBe(0);
    expect(out.brs.tiedOut).toBe(true);
  });

  it('gives identical results for a newest-first statement', () => {
    const rev = bankGrid([...bankRows].reverse());
    const a = run();
    const b = run(rev);
    expect(b.result.matches).toHaveLength(a.result.matches.length);
    expect(b.brs.tiedOut).toBe(true);
    expect(b.brs.bankBalance).toBe(a.brs.bankBalance);
  });

  it('blocks on a dropped statement row via balance validation and localizes it', () => {
    const dropped = bankGrid(bankRows.filter((_, i) => i !== 2));
    const out = run(dropped);
    expect(out.bankValidation.ok).toBe(false);
    expect(out.bankValidation.errors.map((e) => e.code)).toEqual(
      expect.arrayContaining(['BALANCE_TIE_OUT', 'BALANCE_CONTINUITY']),
    );
    const cont = out.bankValidation.errors.find((e) => e.code === 'BALANCE_CONTINUITY')!;
    expect(cont.row).toBe(8); // header row 4 + opening row + the row after the gap
  });

  it('does not tie out (and says so) when the books opening balance is off — e.g. prior-period items', () => {
    const books = booksGrid();
    books[2] = ['Opening Balance', '', '', '', 90000, ''];
    const out = run(bankGrid(), books);
    expect(out.brs.tiedOut).toBe(false);
    expect(out.brs.difference).toBe(-10000);
    expect(out.brs.notes.join(' ')).toMatch(/does not tie out/i);
  });

  it('reports a wrong period instead of an empty reconciliation', () => {
    const out = runBankReconEngine({
      books: { sheetName: 'Ledger', data: booksGrid() },
      bank: { sheetName: 'Stmt', data: bankGrid() },
      periodStart: '2024-01-01',
      periodEnd: '2024-01-31',
    });
    expect(out.periodZeroMessage).toMatch(/No ledger rows/);
  });

  it('swaps sheets that were handed over the wrong way round', () => {
    const out = runBankReconEngine({
      books: { sheetName: 'Stmt', data: bankGrid() },
      bank: { sheetName: 'Ledger', data: booksGrid() },
      periodStart: '2025-04-01',
      periodEnd: '2025-04-30',
    });
    expect(out.notes.join(' ')).toMatch(/that way round/);
    expect(out.brs.tiedOut).toBe(true);
  });

  it('flags an account-number mismatch on the statement header', () => {
    const out = runBankReconEngine({
      books: { sheetName: 'Ledger', data: booksGrid() },
      bank: { sheetName: 'Stmt', data: bankGrid() },
      accountNumber: '9999999999999999',
    });
    expect(out.bankValidation.errors.map((e) => e.code)).toContain('ACCOUNT_MISMATCH');
  });
});

describe('bankMatch', () => {
  let n = 0;
  const tx = (
    source: 'BOOKS' | 'BANK',
    date: string,
    amount: number,
    direction: 'IN' | 'OUT',
    narration = '',
    extra: Partial<BankTxn> = {},
  ): BankTxn => ({
    id: `${source}${++n}`,
    source,
    sourceFormat: 'XLSX',
    txnDate: date,
    amount,
    direction,
    narration,
    parseConfidence: 1,
    sourceLocation: { sheet: source, row: n },
    ...extra,
  });

  it('never matches opposite directions with the same amount', () => {
    const r = bankMatch({
      books: [tx('BOOKS', '2025-04-02', 100, 'IN')],
      bank: [tx('BANK', '2025-04-02', 100, 'OUT')],
      periodEnd: '2025-04-30',
    });
    expect(r.matches).toHaveLength(0);
  });

  it('resolves identical amounts by closest date + narration instead of first-come', () => {
    const b1 = tx('BOOKS', '2025-04-02', 500, 'OUT', 'Alpha Traders');
    const b2 = tx('BOOKS', '2025-04-03', 500, 'OUT', 'Beta Stores');
    const k1 = tx('BANK', '2025-04-03', 500, 'OUT', 'NEFT BETA STORES');
    const k2 = tx('BANK', '2025-04-04', 500, 'OUT', 'NEFT ALPHA TRADERS');
    const r = bankMatch({ books: [b1, b2], bank: [k1, k2], periodEnd: '2025-04-30' });
    expect(r.matches).toHaveLength(2);
    const pair = (b: BankTxn) => r.matches.find((m) => m.books[0].id === b.id)!.bank[0].id;
    expect(pair(b1)).toBe(k2.id);
    expect(pair(b2)).toBe(k1.id);
  });

  it('does not guess between equally plausible combinations in a group match', () => {
    const books = [
      tx('BOOKS', '2025-04-02', 1000, 'OUT', 'Vendor'),
      tx('BOOKS', '2025-04-02', 2000, 'OUT', 'Vendor'),
      tx('BOOKS', '2025-04-02', 3000, 'OUT', 'Vendor'),
    ];
    const bank = [tx('BANK', '2025-04-03', 6000, 'OUT', 'NEFT VENDOR')];
    const r = bankMatch({ books, bank, periodEnd: '2025-04-30' });
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0]).toMatchObject({ kind: 'GROUP', pass: 5 });
    expect(r.matches[0].books).toHaveLength(3);

    const ambiguous = bankMatch({
      books: [
        tx('BOOKS', '2025-04-02', 1000, 'OUT'),
        tx('BOOKS', '2025-04-02', 1000, 'OUT'),
        tx('BOOKS', '2025-04-02', 1000, 'OUT'),
      ],
      bank: [tx('BANK', '2025-04-03', 2000, 'OUT')],
      periodEnd: '2025-04-30',
    });
    expect(ambiguous.matches).toHaveLength(0);
  });

  it('pairs a failed-UPI debit with its refund so neither shows as unmatched', () => {
    const r = bankMatch({
      books: [],
      bank: [
        tx('BANK', '2025-04-02', 750, 'OUT', 'UPI/111/SHOP'),
        tx('BANK', '2025-04-03', 750, 'IN', 'UPI REVERSAL FAILED TXN'),
      ],
      periodEnd: '2025-04-30',
    });
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0].kind).toBe('REVERSAL');
    expect(r.unmatchedBank).toHaveLength(0);
  });

  it('explains amount mismatches on the same cheque number', () => {
    const r = bankMatch({
      books: [tx('BOOKS', '2025-04-02', 1000, 'OUT', 'cheque', { instrumentNo: '456' })],
      bank: [tx('BANK', '2025-04-04', 900, 'OUT', 'CHQ 456', { instrumentNo: '456' })],
      periodEnd: '2025-04-30',
    });
    expect(r.unmatchedBooks[0].reason).toBe('amount_mismatch');
  });

  it('flags duplicates in books and stale cheques', () => {
    const d1 = tx('BOOKS', '2025-04-02', 800, 'OUT', 'Rent');
    const d2 = tx('BOOKS', '2025-04-02', 800, 'OUT', 'Rent');
    const stale = tx('BOOKS', '2025-01-02', 9000, 'OUT', 'cheque issued', { instrumentNo: '99999' });
    const k = tx('BANK', '2025-04-02', 800, 'OUT', 'RENT');
    const r = bankMatch({ books: [d1, d2, stale], bank: [k], periodEnd: '2025-04-30' });
    expect(r.unmatchedBooks.find((u) => u.txn.id === stale.id)!.reason).toBe('stale_cheque');
    expect(r.unmatchedBooks.filter((u) => u.reason === 'duplicate_in_books')).toHaveLength(1);
  });
});

describe('mapper', () => {
  it('produces CREATE_SHEET + WRITE_TABLE for every layout, named distinctly', () => {
    const out = runBankReconEngine({
      books: { sheetName: 'Ledger', data: booksGrid() },
      bank: { sheetName: 'Stmt', data: bankGrid() },
      periodStart: '2025-04-01',
      periodEnd: '2025-04-30',
    });
    const model = {
      result: out.result,
      brs: out.brs,
      periodLabel: 'April 2025',
      periodStart: out.periodStart,
      periodEnd: out.periodEnd,
      runAt: 'now',
      booksSheet: 'Ledger',
      bankSheet: 'Stmt',
    };
    const names = new Set<string>();
    for (const layout of ['categorized', 'books_flat', 'bank_flat', 'matched'] as const) {
      const actions = mapBankReconToActions(layout, model, `Sheet ${layout}`);
      expect(actions[0]).toMatchObject({ type: 'CREATE_SHEET' });
      expect(actions[1]).toMatchObject({ type: 'WRITE_TABLE' });
      names.add((actions[0] as { sheetName: string }).sheetName);
    }
    expect(names.size).toBe(4);
    const cat = mapBankReconToActions('categorized', model, 'S');
    const table = cat[1] as { rows: unknown[][] };
    expect(JSON.stringify(table.rows)).toContain('ties out to 0.00');
  });
});
