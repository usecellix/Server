import { SheetActionPayload } from '../excel-ai/types/sheet-actions.types';
import { NormalizedBankStatementRow } from '../domain-tools/ingestion/bank-statement-parser';
import { isoDateToExcelSerial } from '../domain-tools/ingestion/statement-values';

/**
 * The fixed sheet layout for an imported statement
 * (Root/ATTACHMENT_EXTRACTION_PLAN.md §6.2). Header on row 1, one transaction
 * per row from row 2, nothing else on the sheet, so the result is a clean
 * table the agent and `bank_recon` can read without guessing where it starts.
 */
export const BANK_STATEMENT_HEADERS = [
  'Date',
  'Value Date',
  'Description',
  'Ref No',
  'Debit',
  'Credit',
  'Balance',
  'Source',
  'Flag',
] as const;

const COLUMN_WIDTHS = [95, 95, 360, 120, 110, 110, 120, 80, 300];
const NAVY_FILL = '#203764';
const WHITE = '#FFFFFF';
const DATE_NUMBER_FORMAT = 'dd-mmm-yyyy';
const AMOUNT_NUMBER_FORMAT = '#,##0.00';
const TEXT_NUMBER_FORMAT = '@';

/**
 * Statement text is written into the workbook verbatim, and Excel treats a
 * value starting with one of these as a formula. A leading apostrophe makes it
 * literal text, which is what a description or reference always is.
 */
const FORMULA_TRIGGER_RE = /^[=+\-@\t\r]/;

function literalText(value: string | undefined): string {
  const text = value ?? '';
  return FORMULA_TRIGGER_RE.test(text) ? `'${text}` : text;
}

/** A real Excel date serial, so the cell sorts and filters as a date in any locale. */
function dateCell(iso: string | undefined): number | '' {
  if (!iso) return '';
  return isoDateToExcelSerial(iso) ?? '';
}

export function bankStatementSheetRow(row: NormalizedBankStatementRow): unknown[] {
  return [
    dateCell(row.date),
    dateCell(row.valueDate),
    literalText(row.description),
    literalText(row.refNo),
    row.type === 'debit' ? row.amount : '',
    row.type === 'credit' ? row.amount : '',
    row.balance ?? '',
    literalText(String(row.sourceRowRef.rowOrLine)),
    literalText(row.flags.join('; ')),
  ];
}

export function mapBankStatementToActions(params: {
  sheetName: string;
  rows: NormalizedBankStatementRow[];
  /** Place the new sheet after this one. Omit to let Excel add it at the end. */
  relativeTo?: string;
}): SheetActionPayload[] {
  const { sheetName, rows, relativeTo } = params;
  const lastRow = rows.length + 1;
  const format = (range: string, spec: SheetActionPayload['format']): SheetActionPayload => ({
    type: 'FORMAT_RANGE',
    sheetName,
    range,
    format: spec,
  });

  return [
    {
      type: 'CREATE_SHEET',
      sheetName,
      name: sheetName,
      ...(relativeTo ? { relativeTo, position: 'after' as const } : {}),
    },
    // Text columns are set to Text BEFORE the write. Formatting afterwards is
    // too late: Excel has already turned a reference such as "004521" into the
    // number 4521 by then.
    format(`C2:D${lastRow}`, { numberFormat: TEXT_NUMBER_FORMAT }),
    format(`H2:I${lastRow}`, { numberFormat: TEXT_NUMBER_FORMAT }),
    {
      type: 'WRITE_TABLE',
      sheetName,
      headers: [...BANK_STATEMENT_HEADERS],
      rows: rows.map(bankStatementSheetRow),
    },
    format('A1:I1', { bold: true, fontColor: WHITE, fillColor: NAVY_FILL, horizontalAlignment: 'center' }),
    format(`A2:B${lastRow}`, { numberFormat: DATE_NUMBER_FORMAT, horizontalAlignment: 'center' }),
    format(`E2:G${lastRow}`, { numberFormat: AMOUNT_NUMBER_FORMAT }),
    ...COLUMN_WIDTHS.map(
      (width, col): SheetActionPayload => ({ type: 'SET_COLUMN_WIDTH', sheetName, col, width }),
    ),
    { type: 'FREEZE_PANES', sheetName, freezeRows: 1 },
  ];
}

const BASE_SHEET_NAME = 'Bank Statement';

/**
 * "Bank Statement", or the first of "Bank Statement 2", "Bank Statement 3"…
 * that the workbook does not already have. An import never targets an existing
 * sheet. Excel compares sheet names without regard to case.
 */
export function pickBankStatementSheetName(existingSheetNames: string[] = []): string {
  const taken = new Set(existingSheetNames.map((name) => name.trim().toLowerCase()));
  if (!taken.has(BASE_SHEET_NAME.toLowerCase())) return BASE_SHEET_NAME;
  for (let n = 2; ; n++) {
    const candidate = `${BASE_SHEET_NAME} ${n}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}
