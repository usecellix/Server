import { runBankReconEngine } from './bank-recon.engine';

/**
 * Row-by-row check against the CA-designed answer key for the Sep 2026 demo data
 * (bank_recon_testdata/expected/ground_truth.json). Rows are keyed by date / direction /
 * amount (+ running balance for bank rows), independent of how the engine ids them.
 */
// eslint-disable-next-line @typescript-eslint/no-require-imports
const fx = require('./__fixtures__/demo-sep2026.json') as Record<string, unknown[][]>;
// eslint-disable-next-line @typescript-eslint/no-require-imports
const gt = require('./__fixtures__/ground-truth-sep2026.json');

const k = (t: any) => `${t.txnDate ?? t.date}|${t.direction}|${t.amount}`;
const kb = (t: any) => `${k(t)}|${t.runningBalance ?? t.running_balance}`;
const lk = (t: any) => 'L' + k(t);
const bk = (t: any) => 'B' + kb(t);

const run = () =>
  runBankReconEngine({
    books: { sheetName: 'Bank Ledger', data: fx.sheet1 },
    bank: { sheetName: 'Bank Statement', data: fx.sheet2 },
    priorBrs: { sheetName: 'Prior BRS (Aug-26)', data: fx.sheet3 },
    periodStart: '2026-09-01',
    periodEnd: '2026-09-30',
  });

const gb = new Map<string, any>(gt.bank.map((b: any) => [b.id, b]));
const gl = new Map<string, any>(gt.books.map((b: any) => [b.id, b]));

describe('answer key (Sep 2026 demo data)', () => {
  it('every one of the 30 unreconciled items gets the expected category', () => {
    const o = run();
    const engine = new Map<string, string>();
    o.result.unmatchedBooks.forEach((u) => engine.set('L' + k(u.txn), u.reason));
    o.result.unmatchedBank.forEach((u) => engine.set(bk(u.txn), u.reason));

    const wrong: string[] = [];
    for (const u of gt.unreconciled_items) {
      const key = u.side === 'BANK' ? bk(gb.get(u.id)) : 'L' + k(u);
      if (engine.get(key) !== u.category) wrong.push(`${u.id} ${u.amount}: expected ${u.category}, got ${engine.get(key)}`);
      engine.delete(key);
    }
    expect(wrong).toEqual([]);
    expect([...engine]).toEqual([]); // no unreconciled item the key doesn't know about
  });

  it('finds all 43 expected matches with exactly the expected members (plus the 2 carried-forward ones)', () => {
    const o = run();
    const engine = o.result.matches.map((m) => ({
      books: m.books.map(lk).sort(),
      bank: m.bank.map(bk).sort(),
      pass: m.kind === 'GROUP' ? 'P5' : `P${m.pass}`,
      carry: m.books.some((b) => b.voucherType === 'Carry-forward'),
    }));
    const used = new Set<number>();
    const missing: string[] = [];
    for (const m of gt.matches) {
      const books = (m.books_ids as string[]).map((i) => lk(gl.get(i))).sort();
      const bank = (m.bank_ids as string[]).map((i) => bk(gb.get(i))).sort();
      const idx = engine.findIndex((e, i) => !used.has(i) && JSON.stringify(e.books) === JSON.stringify(books) && JSON.stringify(e.bank) === JSON.stringify(bank));
      if (idx < 0) missing.push(`${m.match_id} ${m.scenario}`);
      else used.add(idx);
    }
    expect(missing).toEqual([]);
    const extra = engine.filter((_, i) => !used.has(i));
    expect(extra).toHaveLength(2);
    expect(extra.every((e) => e.carry)).toBe(true);
  });

  it('reconciles the BRS to the key (books 28,10,451.75 / bank 24,78,981.05 / tied)', () => {
    const o = run();
    expect(o.brs.tiedOut).toBe(true);
    expect(o.brs.booksBalance).toBe(2810451.75);
    expect(o.brs.bankBalance).toBe(2478981.05);
  });
});

describe('statement imported from an attached file (CSV / XLSX / PDF import sheet)', () => {
  // The import feature writes: Date, Value Date, Description, Ref No, Debit, Credit, Balance, Source, Flag —
  // real date cells, header on row 1, no account-detail block, no opening-balance row.
  const serial = (iso: string) => Math.round(Date.parse(iso + 'T00:00:00Z') / 86400000) + 25569;
  const importedGrid = (): unknown[][] => [
    ['Date', 'Value Date', 'Description', 'Ref No', 'Debit', 'Credit', 'Balance', 'Source', 'Flag'],
    ...gt.bank.map((b: any) => [
      serial(b.date),
      b.value_date ? serial(b.value_date) : '',
      b.narration,
      b.reference || b.instrument_no || '',
      b.direction === 'OUT' ? b.amount : '',
      b.direction === 'IN' ? b.amount : '',
      b.running_balance,
      'csv',
      '',
    ]),
  ];

  it('reconciles to the same result as the workbook statement sheet', () => {
    const viaSheet = run();
    const viaImport = runBankReconEngine({
      books: { sheetName: 'Bank Ledger', data: fx.sheet1 },
      bank: { sheetName: 'Bank Statement 2', data: importedGrid() },
      priorBrs: { sheetName: 'Prior BRS (Aug-26)', data: fx.sheet3 },
      periodStart: '2026-09-01',
      periodEnd: '2026-09-30',
    });
    expect(viaImport.bankValidation.errors).toEqual([]);
    expect(viaImport.bankParsed.txns).toHaveLength(65);
    expect(viaImport.result.matches).toHaveLength(viaSheet.result.matches.length);
    expect(viaImport.result.unmatchedBooks).toHaveLength(10);
    expect(viaImport.result.unmatchedBank).toHaveLength(20);
    expect(viaImport.brs.tiedOut).toBe(true);
    expect(viaImport.brs.bankBalance).toBe(2478981.05);
  });
});
