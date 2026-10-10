import { BadRequestException, Injectable, Logger, UnprocessableEntityException } from '@nestjs/common';
import {
  BankStatementParseError,
  BankStatementSkippedRow,
  parseBankStatement,
} from '../domain-tools/ingestion/bank-statement-parser';
import {
  BankStatementVerification,
  verifyBankStatement,
} from '../domain-tools/ingestion/bank-statement-verifier';
import { RawTable, RawTableSource } from '../domain-tools/ingestion/raw-table.types';
import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';
import { ImportBankStatementDto } from './bank-statement.dto';
import {
  mapBankStatementToActions,
  pickBankStatementSheetName,
} from './bank-statement-to-actions.mapper';

const SOURCES: RawTableSource[] = ['pdf', 'xlsx', 'xls', 'csv'];
/** Bounds on the decoded file. Generous for a multi-year statement, small enough to refuse a wrong file early. */
const MAX_ROWS = 120_000;
const MAX_CELLS_PER_ROW = 80;
const MAX_CELL_CHARS = 2_000;
const MAX_TRANSACTIONS = 30_000;
/** How many flagged rows and unread lines the response lists. The sheet's Flag column has all of them. */
const MAX_LISTED = 50;

export interface BankStatementException {
  /** Row on the new sheet, 1-based, so the user can go straight to it. */
  sheetRow: number;
  /** Where the row came from in the file, e.g. "p3 l12". */
  source: string;
  reason: string;
}

export interface BankStatementImportResponse {
  outputSheetName: string;
  statement: {
    fileName: string;
    source: RawTableSource;
    accountNumberMasked?: string;
    transactionCount: number;
    /** ISO dates of the earliest and latest transaction. */
    periodStart?: string;
    periodEnd?: string;
  };
  verification: BankStatementVerification;
  /** Flagged rows, first `MAX_LISTED`. `verification.rowsFlagged` is the full count. */
  exceptions: BankStatementException[];
  /** Lines of the file that were not imported, first `MAX_LISTED`. */
  skipped: BankStatementSkippedRow[];
  skippedCount: number;
  actions: SheetActionPayload[];
}

/**
 * Structural check on the decoded file. Throws a 400: a table in the wrong
 * shape is a client bug, not something the user can act on.
 */
export function assertRawTable(table: unknown): asserts table is RawTable {
  const fail = (reason: string): never => {
    throw new BadRequestException(`Invalid rawTable: ${reason}`);
  };
  const t = table as Partial<RawTable> | null;
  if (!t || typeof t !== 'object') fail('not an object');
  const candidate = t as Partial<RawTable>;
  if (!SOURCES.includes(candidate.source as RawTableSource)) fail('unknown source');
  if (candidate.layout !== 'grid' && candidate.layout !== 'positioned') fail('unknown layout');
  if (typeof candidate.fileName !== 'string' || candidate.fileName.length > 260) fail('bad fileName');
  if (!Array.isArray(candidate.rows)) fail('rows is not an array');
  const rows = candidate.rows as unknown[];
  if (rows.length > MAX_ROWS) fail(`more than ${MAX_ROWS} rows`);

  const finiteOrAbsent = (value: unknown) => value === undefined || Number.isFinite(value);
  for (const raw of rows) {
    const row = raw as { ref?: unknown; page?: unknown; y?: unknown; cells?: unknown };
    if (!row || typeof row.ref !== 'string' || row.ref.length > 40) fail('row without a ref');
    if (!finiteOrAbsent(row.page) || !finiteOrAbsent(row.y)) fail('row with a bad position');
    if (!Array.isArray(row.cells) || row.cells.length > MAX_CELLS_PER_ROW) fail('row with bad cells');
    for (const rawCell of row.cells as unknown[]) {
      const cell = rawCell as { t?: unknown; x0?: unknown; x1?: unknown };
      if (!cell || typeof cell.t !== 'string' || cell.t.length > MAX_CELL_CHARS) fail('cell without text');
      if (!finiteOrAbsent(cell.x0) || !finiteOrAbsent(cell.x1)) fail('cell with a bad position');
    }
  }
}

@Injectable()
export class BankStatementService {
  private readonly logger = new Logger(BankStatementService.name);

  /** The reason a file was refused and its size. Never a cell, never the file name. */
  private refused(code: string, table: RawTable): void {
    this.logger.warn(
      `Bank statement refused: code=${code} source=${table.source} layout=${table.layout} rows=${table.rows.length}`,
    );
  }

  /**
   * Decoded statement in, sheet actions out. No model is called anywhere on
   * this path, and nothing is written: the caller previews the actions and the
   * user accepts them, as with any other change.
   */
  importStatement(request: ImportBankStatementDto): BankStatementImportResponse {
    assertRawTable(request.rawTable);
    const table = request.rawTable;

    let parsed;
    try {
      parsed = parseBankStatement(table);
    } catch (error) {
      if (error instanceof BankStatementParseError) {
        this.refused(error.code, table);
        throw new UnprocessableEntityException({ code: error.code, message: error.message });
      }
      throw error;
    }
    if (parsed.rows.length > MAX_TRANSACTIONS) {
      this.refused('too_many_transactions', table);
      throw new UnprocessableEntityException({
        code: 'too_many_transactions',
        message: `This statement has ${parsed.rows.length} transactions. The most that can be imported at once is ${MAX_TRANSACTIONS}. Split it into shorter periods.`,
      });
    }

    const { rows, verification } = verifyBankStatement(parsed);
    const outputSheetName = pickBankStatementSheetName(request.existingSheetNames);
    const actions = mapBankStatementToActions({
      sheetName: outputSheetName,
      rows,
      relativeTo: request.activeSheetName,
    });

    const exceptions: BankStatementException[] = [];
    rows.forEach((row, index) => {
      if (!row.flags.length || exceptions.length >= MAX_LISTED) return;
      exceptions.push({
        sheetRow: index + 2,
        source: String(row.sourceRowRef.rowOrLine),
        reason: row.flags.join('; '),
      });
    });

    const dates = rows.map((row) => row.date).filter(Boolean).sort();

    // Counts only. Statement contents never go to the log.
    this.logger.log(
      `Bank statement parsed: source=${table.source} transactions=${rows.length} state=${verification.state} ` +
        `failed=${verification.rowsFailed} flagged=${verification.rowsFlagged} skipped=${parsed.skipped.length}`,
    );

    return {
      outputSheetName,
      statement: {
        fileName: table.fileName,
        source: table.source,
        ...(parsed.accountNumberMasked ? { accountNumberMasked: parsed.accountNumberMasked } : {}),
        transactionCount: rows.length,
        ...(dates.length ? { periodStart: dates[0], periodEnd: dates[dates.length - 1] } : {}),
      },
      verification,
      exceptions,
      skipped: parsed.skipped.slice(0, MAX_LISTED),
      skippedCount: parsed.skipped.length,
      actions,
    };
  }
}
