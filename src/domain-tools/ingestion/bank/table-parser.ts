import { BankTxn, Direction, ParsedTable, SourceFormat, TxnSource } from '../../../bank-recon/types';
import { parseNarration, stripLeadingZeros } from '../../bank/narration-parser';
import { cleanText, normHeader, parseBankAmount, parseBankDate, roundMoney } from './bank-utils';

/**
 * One tabular parser for BOTH the bank statement and the books bank ledger — they differ
 * only in how Debit/Credit map to money-in/out, which `role` decides. Excel and CSV grids
 * (and, later, PDF-extracted tables) all arrive here as `unknown[][]`.
 */

export type BankField =
  | 'valueDate'
  | 'date'
  | 'instrument'
  | 'reference'
  | 'drcr'
  | 'debit'
  | 'credit'
  | 'receipt'
  | 'payment'
  | 'amount'
  | 'balance'
  | 'vchType'
  | 'vchNo'
  | 'narration';

const FIELD_RULES: Array<[BankField, RegExp]> = [
  ['valueDate', /value ?(date|dt)/],
  ['vchType', /^(vch|voucher) ?type$/],
  ['vchNo', /^(vch|voucher) ?(no|num|number)\b/],
  ['instrument', /^(chq|cheque|check|instrument|instr)\b.*(no|num|number)|^(chq|cheque|instrument)$/],
  ['reference', /^(ref|reference|utr|txn ref|transaction ref|txn id|transaction id|bank ref)\b/],
  ['drcr', /^(dr ?\/ ?cr|cr ?\/ ?dr|dr cr|cr dr|dr or cr|txn type|transaction type|type)$/],
  ['balance', /balance|^bal\b/],
  ['debit', /^(debit|withdrawal|withdrawals|dr|paid out)\b/],
  ['credit', /^(credit|deposit|deposits|cr|paid in)\b/],
  ['receipt', /^receipts?\b/],
  ['payment', /^payments?\b/],
  ['amount', /^(txn |tran |transaction )?amount( in rs)?$|^amt$/],
  ['date', /^(txn |tran |trans |transaction |posting |post |voucher |vch |book |booking )?(date|dt)$/],
  ['narration', /narration|description|particulars|remarks|details/],
];

export function classifyHeader(raw: unknown): BankField | null {
  const h = normHeader(raw);
  if (!h || h.length > 40) return null;
  for (const [field, re] of FIELD_RULES) {
    if (re.test(h)) return field;
  }
  return null;
}

export interface HeaderMap {
  headerRow: number; // 0-based
  cols: Partial<Record<BankField, number>>;
  narrationCols: number[];
  particularsCol?: number;
  score: number;
}

const AMOUNT_FIELDS: BankField[] = ['debit', 'credit', 'receipt', 'payment', 'amount'];

/** Finds the real header row below any bank-name / address / account-detail junk. */
export function detectHeader(grid: unknown[][], preferredRow?: number): HeaderMap | null {
  const limit = Math.min(grid.length, 40);
  let best: HeaderMap | null = null;
  const rows = preferredRow && preferredRow > 0 ? [preferredRow - 1] : Array.from({ length: limit }, (_, i) => i);
  for (const i of rows) {
    const row = grid[i];
    if (!row) continue;
    const cols: Partial<Record<BankField, number>> = {};
    const narrationCols: number[] = [];
    let particularsCol: number | undefined;
    row.forEach((cell, c) => {
      const f = classifyHeader(cell);
      if (!f) return;
      if (f === 'narration') {
        narrationCols.push(c);
        if (normHeader(cell) === 'particulars') particularsCol = c;
      }
      if (cols[f] === undefined) cols[f] = c;
    });
    const fields = Object.keys(cols) as BankField[];
    const hasDate = cols.date !== undefined;
    const hasAmount = AMOUNT_FIELDS.some((f) => cols[f] !== undefined);
    if (!hasDate || !hasAmount) continue;
    const score = fields.length;
    if (!best || score > best.score) {
      best = { headerRow: i, cols, narrationCols, particularsCol, score };
    }
  }
  return best && best.score >= 3 ? best : null;
}

interface RawCandidate {
  date: string;
  valueDate: string;
  rowNumber: number; // 1-based sheet row
  amount: number;
  dir: Direction | null; // null → decide from balance / sign later
  signHint: Direction;
  balance?: number;
  narration: string;
  counterparty?: string;
  instrumentRaw: string;
  referenceRaw: string;
  vchType?: string;
  vchNo?: string;
  confidence: number;
}

const OPENING_RE = /opening\s*bal|balance\s*b\/?f|brought\s*forward|^b\/f\b|bal\s*b\/f/i;
const CLOSING_RE = /closing\s*bal|balance\s*c\/?f|carried\s*forward|^c\/f\b|bal\s*c\/f|balance\s*c\/d/i;
const TOTAL_RE = /^(grand\s+)?total\b|page\s+total|^sub\s*total|^continued|statement\s+summary|^end of statement|^\*+/i;

function cellStr(row: unknown[], idx: number | undefined): string {
  if (idx === undefined) return '';
  return cleanText(row[idx]);
}

function signedBalance(raw: unknown, role: TxnSource): number | undefined {
  const p = parseBankAmount(raw);
  if (!p.present) return undefined;
  // Bank: CR = positive (money held). Books (Tally): Dr = positive (asset).
  if (p.suffix === 'DR') return role === 'BANK' ? -Math.abs(p.value) : Math.abs(p.value);
  if (p.suffix === 'CR') return role === 'BANK' ? Math.abs(p.value) : -Math.abs(p.value);
  return p.value;
}

function dirFromFlag(flag: string, role: TxnSource): Direction | null {
  const f = flag.trim().toUpperCase();
  if (!f) return null;
  const isCredit = /^(CR|C|CREDIT|DEPOSIT|RECEIPT)\b/.test(f);
  const isDebit = /^(DR|D|DEBIT|WITHDRAWAL|PAYMENT)\b/.test(f);
  if (!isCredit && !isDebit) return null;
  if (role === 'BANK') return isCredit ? 'IN' : 'OUT';
  return isDebit ? 'IN' : 'OUT';
}

export interface ParseTableOptions {
  role: TxnSource;
  sheetName: string;
  sourceFormat?: SourceFormat;
  /** 1-based header row override. */
  headersRow?: number;
}

export function parseBankTable(grid: unknown[][], opts: ParseTableOptions): ParsedTable {
  const { role, sheetName } = opts;
  const sourceFormat: SourceFormat = opts.sourceFormat ?? 'XLSX';
  const warnings: string[] = [];
  const empty: ParsedTable = {
    txns: [],
    openingBalance: null,
    closingBalance: null,
    headerRow: 0,
    preambleText: [],
    warnings,
  };

  const header = detectHeader(grid, opts.headersRow) ?? detectHeader(grid);
  if (!header) {
    warnings.push(
      `Could not find a header row with a Date column and Debit/Credit/Amount column(s) in "${sheetName}".`,
    );
    return empty;
  }
  const { cols } = header;
  const preambleText: string[] = [];
  for (let i = 0; i < header.headerRow; i++) {
    const t = (grid[i] ?? []).map((c) => cleanText(c)).filter(Boolean).join(' ');
    if (t) preambleText.push(t);
  }

  const hasSplitAmounts =
    cols.debit !== undefined || cols.credit !== undefined || cols.receipt !== undefined || cols.payment !== undefined;

  let opening: number | null = null;
  let closing: number | null = null;
  const cands: RawCandidate[] = [];
  let lastDate = '';

  const rowLooksLikeHeader = (row: unknown[]): boolean => {
    let hits = 0;
    for (const c of row) if (typeof c === 'string' && classifyHeader(c)) hits += 1;
    return hits >= 3;
  };

  for (let i = header.headerRow + 1; i < grid.length; i++) {
    const row = grid[i] ?? [];
    if (!row.some((c) => cleanText(c) !== '')) continue;
    if (rowLooksLikeHeader(row)) continue; // repeated page header

    const rowText = row.map((c) => cleanText(c)).filter(Boolean).join(' ');
    const date = parseBankDate(row[cols.date!]);
    const narrationParts = header.narrationCols.map((c) => cellStr(row, c)).filter(Boolean);
    const narration = narrationParts.join(' ').trim();

    // Amount extraction
    let debit = 0;
    let credit = 0;
    let receipt = 0;
    let payment = 0;
    let single = 0;
    let singleSuffix: 'CR' | 'DR' | undefined;
    if (cols.debit !== undefined) debit = parseBankAmount(row[cols.debit]).value;
    if (cols.credit !== undefined) credit = parseBankAmount(row[cols.credit]).value;
    if (cols.receipt !== undefined) receipt = parseBankAmount(row[cols.receipt]).value;
    if (cols.payment !== undefined) payment = parseBankAmount(row[cols.payment]).value;
    if (cols.amount !== undefined) {
      const p = parseBankAmount(row[cols.amount]);
      single = p.value;
      singleSuffix = p.suffix;
    }
    const balance = cols.balance !== undefined ? signedBalance(row[cols.balance], role) : undefined;

    const anyAmount = debit !== 0 || credit !== 0 || receipt !== 0 || payment !== 0 || single !== 0;

    // Opening / closing balance rows — never transactions.
    if (!date || !anyAmount || OPENING_RE.test(narration) || CLOSING_RE.test(narration)) {
      if (OPENING_RE.test(rowText) || OPENING_RE.test(narration)) {
        if (balance !== undefined) opening = balance;
        else if (anyAmount) opening = openingFromAmounts(role, debit, credit, receipt, payment, single, singleSuffix);
        continue;
      }
      if (CLOSING_RE.test(rowText) || CLOSING_RE.test(narration)) {
        if (balance !== undefined) closing = balance;
        else if (anyAmount) {
          // A Tally ledger prints its closing balance on the side OPPOSITE to the balance
          // (a debit balance sits in the Credit column), so the sign flips for books.
          const v = openingFromAmounts(role, debit, credit, receipt, payment, single, singleSuffix);
          closing = role === 'BOOKS' && (debit || credit) ? -v : v;
        }
        continue;
      }
    }
    if (TOTAL_RE.test(narration) || TOTAL_RE.test(rowText)) {
      if (!date) continue;
    }

    if (!anyAmount) {
      // Continuation line: no date, no amount, free text → belongs to the txn above.
      if (!date && narration && cands.length) {
        const prev = cands[cands.length - 1];
        prev.narration = `${prev.narration} ${narration}`.trim();
      }
      continue;
    }

    // Money row. Never drop a row that has an amount; borrow the previous date when missing.
    let rowDate = date;
    let confidence = 1;
    if (!rowDate) {
      if (!lastDate) {
        warnings.push(`Row ${i + 1} of "${sheetName}" has an amount but no date and was skipped.`);
        continue;
      }
      rowDate = lastDate;
      confidence = 0.8;
      warnings.push(`Row ${i + 1} of "${sheetName}" had no date — used ${lastDate}.`);
    } else {
      lastDate = rowDate;
    }

    const base = {
      date: rowDate,
      valueDate: parseBankDate(cols.valueDate !== undefined ? row[cols.valueDate] : ''),
      rowNumber: i + 1,
      balance,
      narration,
      counterparty:
        role === 'BOOKS' && header.particularsCol !== undefined && header.narrationCols.length > 1
          ? cellStr(row, header.particularsCol)
          : undefined,
      instrumentRaw: cellStr(row, cols.instrument),
      referenceRaw: cellStr(row, cols.reference),
      vchType: cellStr(row, cols.vchType) || undefined,
      vchNo: cellStr(row, cols.vchNo) || undefined,
    };

    if (hasSplitAmounts && cols.amount === undefined) {
      const ins = role === 'BANK' ? credit + receipt : debit + receipt;
      const outs = role === 'BANK' ? debit + payment : credit + payment;
      if (ins > 0) cands.push({ ...base, amount: roundMoney(ins), dir: 'IN', signHint: 'IN', confidence });
      if (outs > 0) cands.push({ ...base, amount: roundMoney(outs), dir: 'OUT', signHint: 'OUT', confidence });
      // Negative values in a debit/credit column are reversals of that column.
      const neg = [debit, credit, receipt, payment].some((v) => v < 0);
      if (ins <= 0 && outs <= 0 && neg) {
        const v = Math.abs(debit || credit || receipt || payment);
        const wasDebitCol = debit < 0 || payment < 0;
        const dir: Direction = wasDebitCol ? (role === 'BANK' ? 'IN' : 'OUT') : role === 'BANK' ? 'OUT' : 'IN';
        cands.push({ ...base, amount: v, dir, signHint: dir, confidence: confidence * 0.8 });
      }
    } else {
      const flag = cols.drcr !== undefined ? cellStr(row, cols.drcr) : '';
      let dir = dirFromFlag(flag, role);
      if (!dir && singleSuffix) dir = dirFromFlag(singleSuffix, role);
      const abs = Math.abs(single);
      cands.push({
        ...base,
        amount: roundMoney(abs),
        dir,
        signHint: single < 0 ? 'OUT' : 'IN',
        confidence: dir ? confidence : confidence * 0.7,
      });
    }
  }

  if (!cands.length) {
    warnings.push(`No transactions found in "${sheetName}" below the header row.`);
    return { ...empty, headerRow: header.headerRow + 1, preambleText, openingBalance: opening, closingBalance: closing };
  }

  // Newest-first exports: walk oldest → newest so running-balance deltas make sense.
  const firstDate = cands[0].date;
  const lastCandDate = cands[cands.length - 1].date;
  const descending = firstDate > lastCandDate;
  const ordered = descending ? [...cands].reverse() : cands;

  // Resolve undecided directions from the running balance.
  let prevBalance: number | undefined = opening ?? undefined;
  for (const c of ordered) {
    if (c.dir === null) {
      let resolved: Direction | null = null;
      if (c.balance !== undefined && prevBalance !== undefined) {
        const delta = roundMoney(c.balance - prevBalance);
        // Balance rises on money in, for both bank and books ledgers.
        if (Math.abs(delta - c.amount) < 0.011) resolved = 'IN';
        else if (Math.abs(delta + c.amount) < 0.011) resolved = 'OUT';
      }
      if (resolved) {
        c.dir = resolved;
        c.confidence = Math.min(c.confidence, 0.85);
      } else {
        c.dir = c.signHint;
        c.confidence = Math.min(c.confidence, 0.7);
      }
    }
    if (c.balance !== undefined) prevBalance = c.balance;
  }

  const txns: BankTxn[] = ordered.map((c) => {
    const parsedNarr = parseNarration(c.narration);
    const { instrumentNo, referenceNo } = splitInstrumentAndRef(c.instrumentRaw, c.referenceRaw);
    const instr = instrumentNo ?? parsedNarr.instrumentNo;
    const ref = referenceNo ?? parsedNarr.referenceNo;
    const counterparty = c.counterparty || parsedNarr.counterparty;
    return {
      id: `${role}:${sheetName}:${c.rowNumber}:${c.dir}`,
      source: role,
      sourceFormat,
      txnDate: c.date,
      valueDate: c.valueDate || undefined,
      amount: c.amount,
      direction: c.dir as Direction,
      instrumentNo: instr,
      referenceNo: ref,
      narration: c.narration,
      counterparty: counterparty || undefined,
      channel: parsedNarr.channel,
      voucherType: c.vchType,
      voucherNo: c.vchNo,
      runningBalance: c.balance,
      parseConfidence: roundMoney(c.confidence * 100) / 100,
      sourceLocation: { sheet: sheetName, row: c.rowNumber },
    };
  });

  // Derive balances when the sheet carries a running-balance column but no explicit rows.
  const signed = (t: BankTxn) => (t.direction === 'IN' ? t.amount : -t.amount);
  if (opening === null) {
    const first = txns.find((t) => t.runningBalance !== undefined);
    if (first && txns[0] === first) opening = roundMoney(first.runningBalance! - signed(first));
  }
  if (closing === null) {
    const lastWithBal = [...txns].reverse().find((t) => t.runningBalance !== undefined);
    if (lastWithBal) closing = lastWithBal.runningBalance!;
  }

  return {
    txns,
    openingBalance: opening,
    closingBalance: closing,
    headerRow: header.headerRow + 1,
    preambleText,
    warnings,
  };
}

function openingFromAmounts(
  role: TxnSource,
  debit: number,
  credit: number,
  receipt: number,
  payment: number,
  single: number,
  suffix?: 'CR' | 'DR',
): number {
  // Books: debit balance positive. Bank: credit balance positive.
  const posSide = role === 'BANK' ? credit + receipt : debit + receipt;
  const negSide = role === 'BANK' ? debit + payment : credit + payment;
  if (posSide || negSide) return roundMoney(posSide - negSide);
  if (suffix === 'DR') return role === 'BANK' ? -Math.abs(single) : Math.abs(single);
  if (suffix === 'CR') return role === 'BANK' ? Math.abs(single) : -Math.abs(single);
  return single;
}

/** A "Chq./Ref.No." column holds either a cheque number or a UTR — tell them apart. */
export function splitInstrumentAndRef(
  instrumentRaw: string,
  referenceRaw: string,
): { instrumentNo?: string; referenceNo?: string } {
  let instrumentNo: string | undefined;
  let referenceNo: string | undefined;
  const classify = (v: string) => {
    const t = v.trim();
    if (!t || /^0+$/.test(t)) return;
    if (/^\d{5,8}$/.test(t)) {
      instrumentNo = instrumentNo ?? stripLeadingZeros(t);
    } else if (/[A-Za-z0-9]{6,}/.test(t)) {
      referenceNo = referenceNo ?? t.toUpperCase().replace(/\s+/g, '');
    }
  };
  classify(instrumentRaw);
  classify(referenceRaw);
  return { instrumentNo, referenceNo };
}
