import { UnprocessableEntityException } from '@nestjs/common';
import {
  BankStatementParseError,
  parseBankStatement,
} from '../domain-tools/ingestion/bank-statement-parser';
import { verifyBankStatement } from '../domain-tools/ingestion/bank-statement-verifier';
import { RawTable } from '../domain-tools/ingestion/raw-table.types';
import { assertRawTable } from '../bank-statement/bank-statement.service';
import {
  BANK_STATEMENT_HEADERS,
  bankStatementSheetRow,
} from '../bank-statement/bank-statement-to-actions.mapper';

export interface AttachedStatement {
  /** Header + one row per transaction, in the same shape the "import statement" sheet uses. */
  grid: unknown[][];
  /** Sheet row (1-based) → where that row sits in the attached file, e.g. "p3 l12" or "row 45". */
  sourceRefs: Map<number, string>;
  openingBalance?: number;
  closingBalance?: number;
  fileName: string;
  source: RawTable['source'];
  transactionCount: number;
  accountNumberMasked?: string;
  /** Importer's own checks (running balance, direction fixes, stated totals). */
  verification: { state: string; notes: string[]; rowsFailed: number; rowsFlagged: number };
}

/**
 * Reads a decoded CSV / Excel / PDF attachment straight into transaction rows, reusing the
 * statement importer's parser and verifier — nothing is written to a sheet. The result is
 * shaped like the importer's output sheet so the reconciliation reads it like any statement.
 */
export function readAttachedStatement(rawTable: unknown): AttachedStatement {
  assertRawTable(rawTable);
  const table = rawTable;

  let parsed;
  try {
    parsed = parseBankStatement(table);
  } catch (error) {
    if (error instanceof BankStatementParseError) {
      throw new UnprocessableEntityException({ code: error.code, message: error.message });
    }
    throw error;
  }

  const { rows, verification } = verifyBankStatement(parsed);
  const sourceRefs = new Map<number, string>();
  const grid: unknown[][] = [[...BANK_STATEMENT_HEADERS]];
  rows.forEach((row, i) => {
    grid.push(bankStatementSheetRow(row));
    sourceRefs.set(i + 2, String(row.sourceRowRef.rowOrLine));
  });

  return {
    grid,
    sourceRefs,
    openingBalance: parsed.openingBalance ?? verification.openingBalance,
    closingBalance: parsed.closingBalance ?? verification.closingBalance,
    fileName: table.fileName,
    source: table.source,
    transactionCount: rows.length,
    accountNumberMasked: parsed.accountNumberMasked,
    verification: {
      state: verification.state,
      notes: verification.notes,
      rowsFailed: verification.rowsFailed,
      rowsFlagged: verification.rowsFlagged,
    },
  };
}
