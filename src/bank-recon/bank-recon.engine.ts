import { parseCarryForward } from '../domain-tools/ingestion/bank/carry-forward-parser';
import { addDays, parseBankDate } from '../domain-tools/ingestion/bank/bank-utils';
import { detectTableRole } from '../domain-tools/ingestion/bank/bank-file-detector';
import { parseBankTable } from '../domain-tools/ingestion/bank/table-parser';
import { dedupeRepeatedRows, validateParsedTable } from '../domain-tools/ingestion/bank/validator';
import { bankMatch } from '../domain-tools/bank/match-passes';
import { balanceAsAt, buildBrs } from './brs-builder';
import {
  BankMatchResult,
  BankMatchSettings,
  BankTxn,
  BrsResult,
  ParsedTable,
  SourceFormat,
  ValidationResult,
} from './types';

export interface EngineSheet {
  sheetName: string;
  data: unknown[][];
  headersRow?: number;
  fileType?: 'XLSX' | 'CSV';
}

export interface EngineInput {
  books: EngineSheet;
  bank: EngineSheet;
  /** The bank statement came from an attached file: sheet row → place in that file. */
  bankSourceRefs?: Map<number, string>;
  /** Previous period's BRS — its uncleared items are injected as opening ledger entries. */
  priorBrs?: EngineSheet;
  periodStart?: string;
  periodEnd?: string;
  booksOpeningBalance?: number;
  bankOpeningBalance?: number;
  accountNumber?: string;
  settings?: Partial<BankMatchSettings>;
}

export interface EngineOutput {
  result: BankMatchResult;
  brs: BrsResult;
  periodStart: string;
  periodEnd: string;
  booksValidation: ValidationResult;
  bankValidation: ValidationResult;
  booksParsed: ParsedTable;
  bankParsed: ParsedTable;
  booksRowCount: number;
  bankRowCount: number;
  /** Set when the supplied period leaves nothing to reconcile on one side. */
  periodZeroMessage?: string;
  carryForwardCount: number;
  notes: string[];
}

const BOUNDARY_DAYS_BEFORE = 10;

const DATE_SRC = String.raw`(\d{1,2}[-/ ][A-Za-z]{3}[a-z]*[-/ ,]+\d{2,4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})`;
const RANGE_RE = new RegExp(DATE_SRC + String.raw`\s*(?:to|-|–|—)\s*` + DATE_SRC, 'i');

/** "Statement Period: 01-Sep-2026 to 30-Sep-2026" style line in the text above a header row. */
function periodFromPreamble(lines: string[]): { start: string; end: string } | null {
  for (const line of lines) {
    const m = line.match(RANGE_RE);
    if (!m) continue;
    const start = parseBankDate(m[1]);
    const end = parseBankDate(m[2]);
    if (start && end && start <= end) return { start, end };
  }
  return null;
}

export function runBankReconEngine(input: EngineInput): EngineOutput {
  const notes: string[] = [];
  let booksSheet = input.books;
  let bankSheet = input.bank;

  // Guard against the two sheets being handed over the wrong way round.
  const rb = detectTableRole(booksSheet.data, booksSheet.sheetName);
  const rk = detectTableRole(bankSheet.data, bankSheet.sheetName);
  if (rb.role === 'BANK' && rk.role === 'BOOKS') {
    [booksSheet, bankSheet] = [bankSheet, booksSheet];
    notes.push(`"${booksSheet.sheetName}" looks like the books ledger and "${bankSheet.sheetName}" like the bank statement — read them that way round.`);
  }

  const fmt = (s: EngineSheet): SourceFormat => (s.fileType === 'CSV' ? 'CSV' : 'XLSX');
  const booksParsed = parseBankTable(booksSheet.data, {
    role: 'BOOKS',
    sheetName: booksSheet.sheetName,
    headersRow: booksSheet.headersRow,
    sourceFormat: fmt(booksSheet),
  });
  const bankParsed = parseBankTable(bankSheet.data, {
    role: 'BANK',
    sheetName: bankSheet.sheetName,
    headersRow: bankSheet.headersRow,
    sourceFormat: fmt(bankSheet),
  });

  if (input.bankSourceRefs) {
    // Attached file: there is no sheet row to jump to, only a place in the file.
    const refs = input.bankSourceRefs;
    for (const t of bankParsed.txns) {
      t.sourceLocation = { ...t.sourceLocation, external: true, ref: refs.get(t.sourceLocation.row) };
    }
  }

  const dedup = dedupeRepeatedRows(bankParsed.txns);
  if (dedup.removed.length) {
    notes.push(`${dedup.removed.length} repeated statement row(s) (overlapping pages) were removed.`);
    bankParsed.txns = dedup.txns;
  }

  const booksAll = booksParsed.txns;
  const bankAll = bankParsed.txns;

  // Period: explicit, else the books ledger's own date range.
  const dates = booksAll.map((t) => t.txnDate).sort();
  const stated =
    periodFromPreamble(booksParsed.preambleText) ?? periodFromPreamble(bankParsed.preambleText);
  const periodStart = input.periodStart ?? stated?.start ?? dates[0] ?? bankAll[0]?.txnDate ?? '';
  const periodEnd =
    input.periodEnd ?? stated?.end ?? dates[dates.length - 1] ?? bankAll[bankAll.length - 1]?.txnDate ?? '';
  const carryForward = input.priorBrs ? parseCarryForward(input.priorBrs.data, input.priorBrs.sheetName) : [];

  const booksValidation = validateParsedTable(booksParsed, {
    label: 'Books ledger',
    periodStart,
    periodEnd,
    openingBalance: input.booksOpeningBalance,
  });
  const bankValidation = validateParsedTable(bankParsed, {
    label: 'Bank statement',
    periodStart,
    periodEnd,
    accountNumber: input.accountNumber,
    openingBalance: input.bankOpeningBalance,
  });

  const empty = (message?: string): EngineOutput => ({
    result: { matches: [], unmatchedBooks: [], unmatchedBank: [] },
    brs: buildBrs({
      result: { matches: [], unmatchedBooks: [], unmatchedBank: [] },
      periodEnd,
      booksBalance: null,
      bankBalance: null,
    }),
    periodStart,
    periodEnd,
    booksValidation,
    bankValidation,
    booksParsed,
    bankParsed,
    booksRowCount: booksAll.length,
    bankRowCount: bankAll.length,
    periodZeroMessage: message,
    carryForwardCount: carryForward.length,
    notes,
  });

  if (!booksAll.length || !bankAll.length) {
    return empty(
      !booksAll.length
        ? `I couldn't read any transactions from "${booksSheet.sheetName}". ${booksParsed.warnings[0] ?? ''}`.trim()
        : `I couldn't read any transactions from "${bankSheet.sheetName}". ${bankParsed.warnings[0] ?? ''}`.trim(),
    );
  }

  const inPeriod = (t: BankTxn) => t.txnDate >= periodStart;
  const books = [...carryForward, ...booksAll.filter(inPeriod)];
  if (input.periodStart && input.periodEnd && !books.some((t) => t.txnDate >= periodStart && t.txnDate <= periodEnd)) {
    return empty(`No ledger rows found between ${periodStart} and ${periodEnd} — check the date column or try a different period.`);
  }
  const clearing = input.settings?.clearingWindowDays ?? 15;
  const lo = addDays(periodStart, -BOUNDARY_DAYS_BEFORE);
  const hi = addDays(periodEnd, clearing + 5);
  const bank = bankAll.filter((t) => t.txnDate >= lo && t.txnDate <= hi);
  if (input.periodStart && input.periodEnd && !bank.some((t) => t.txnDate >= periodStart && t.txnDate <= periodEnd)) {
    return empty(`No bank statement rows found between ${periodStart} and ${periodEnd} — check the statement's date column or period.`);
  }

  const result = bankMatch({ books, bank, settings: input.settings, periodEnd });

  // Balances as at period end. Books opening is rolled forward to the period start first.
  let booksOpening: number | null = input.booksOpeningBalance ?? booksParsed.openingBalance;
  if (booksOpening !== null) {
    booksOpening += booksAll
      .filter((t) => t.txnDate < periodStart)
      .reduce((s, t) => s + (t.direction === 'IN' ? t.amount : -t.amount), 0);
  }
  // Carry-forward items were already booked in the previous period, so they are part of the
  // opening balance — never add them to the ledger movement a second time.
  const booksBalance = balanceAsAt(books.filter((t) => t.voucherType !== 'Carry-forward'), booksOpening, periodEnd);
  const bankBalance = balanceAsAt(bankAll, input.bankOpeningBalance ?? bankParsed.openingBalance, periodEnd);

  const brs = buildBrs({
    result: {
      ...result,
      unmatchedBank: result.unmatchedBank.filter((u) => u.txn.txnDate >= periodStart),
    },
    periodEnd,
    booksBalance,
    bankBalance,
    validationWarnings: [...booksValidation.errors, ...bankValidation.errors].map((e) => e.message),
  });

  return {
    result,
    brs,
    periodStart,
    periodEnd,
    booksValidation,
    bankValidation,
    booksParsed,
    bankParsed,
    booksRowCount: books.length - carryForward.length,
    bankRowCount: bank.length,
    carryForwardCount: carryForward.length,
    notes,
  };
}
