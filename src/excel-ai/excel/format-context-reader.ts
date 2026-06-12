import { ColumnMeta, SheetSnapshot, WorkbookContext } from '../../types/cellix.types';
import { SheetActionPayload } from '../types/sheet-actions.types';

export interface ColumnFormatDescription {
  columnIndex: number;
  header: string;
  type: string;
  numberFormat?: string;
  exampleValues: string[];
  rule: string;
}

export interface SheetFormatContext {
  sheetName: string;
  columns: ColumnFormatDescription[];
  hasHeaders: boolean;
  firstDataRow: number;
  lastDataRow: number;
}

const FORMAT_ACTION_TYPES = new Set(['ADD_ROW', 'INSERT_ROW', 'SET_CELL', 'SET_FORMULA']);

function inferDetectedType(
  sampleValues: (string | number | null)[],
  numberFormat?: string,
): string {
  if (numberFormat) {
    const lower = numberFormat.toLowerCase();
    if (/[dmy]/i.test(lower) && !/^\#/.test(numberFormat)) return 'date';
    if (/[$₹€£]/.test(numberFormat)) return 'currency';
    if (/[\#0]/.test(numberFormat)) return 'number';
  }

  const nonEmpty = sampleValues.filter((v) => v != null && String(v).trim() !== '');
  if (nonEmpty.every((v) => typeof v === 'boolean')) return 'boolean';
  if (nonEmpty.every((v) => typeof v === 'number')) return 'number';
  if (nonEmpty.length > 0) return 'text';
  return 'unknown';
}

function resolveColumnMeta(sheet: SheetSnapshot): ColumnMeta[] {
  if (sheet.columnMeta?.length) {
    return sheet.columnMeta;
  }

  const colCount = Math.max(sheet.colCount, sheet.headers.length, 1);
  return Array.from({ length: colCount }, (_, index) => {
    const header = sheet.headers[index] ?? '';
    const sampleValues = sheet.sampleData
      .map((row) => row[index] ?? null)
      .filter((v) => v != null && String(v).trim() !== '')
      .slice(0, 5);

    return {
      index,
      header,
      sampleValues,
      detectedType: inferDetectedType(sampleValues),
    };
  });
}

function buildColumnDescription(col: ColumnMeta): ColumnFormatDescription {
  const detectedType = col.detectedType ?? inferDetectedType(col.sampleValues, col.numberFormat);
  const exampleValues = col.sampleValues
    .filter((v) => v != null)
    .slice(0, 3)
    .map((v) => String(v));

  let rule = '';

  switch (detectedType) {
    case 'date':
      rule = col.numberFormat
        ? `MUST use Excel numberFormat "${col.numberFormat}" for all dates in this column`
        : 'Column contains dates — use ISO format (YYYY-MM-DD) and apply a date numberFormat';
      break;
    case 'currency':
      rule = col.numberFormat
        ? `Apply currency format "${col.numberFormat}" to all values in this column`
        : 'Column contains currency — use a currency numberFormat';
      break;
    case 'number':
      rule = col.numberFormat
        ? `Apply number format "${col.numberFormat}"`
        : 'Column contains numbers — preserve existing number format';
      break;
    case 'boolean':
      rule = 'Column contains boolean values (TRUE/FALSE)';
      break;
    default:
      rule = 'Column contains text';
  }

  return {
    columnIndex: col.index,
    header: col.header ?? `Column ${col.index + 1}`,
    type: detectedType,
    numberFormat: col.numberFormat,
    exampleValues,
    rule,
  };
}

export function extractFormatContext(context: WorkbookContext): SheetFormatContext[] {
  return context.sheets.map((sheet) => ({
    sheetName: sheet.sheetName,
    hasHeaders: sheet.headers.length > 0 && sheet.headers.some((h) => h !== ''),
    firstDataRow: sheet.headers.length > 0 ? 2 : 1,
    lastDataRow: sheet.rowCount,
    columns: resolveColumnMeta(sheet)
      .slice()
      .sort((a, b) => a.index - b.index)
      .map((col) => buildColumnDescription(col)),
  }));
}

function primaryDateColumn(sheetFmt: SheetFormatContext): ColumnFormatDescription | undefined {
  return sheetFmt.columns.find((col) => col.type === 'date' && col.numberFormat);
}

function actionSheetName(action: SheetActionPayload, context: WorkbookContext): string {
  return action.sheetName ?? context.activeSheet;
}

/**
 * After parsing AI response, inject missing numberFormats for date columns.
 */
export function injectMissingFormats(
  actions: SheetActionPayload[],
  context: WorkbookContext,
): SheetActionPayload[] {
  const formatContexts = extractFormatContext(context);

  return actions.map((action) => {
    if (!FORMAT_ACTION_TYPES.has(action.type)) {
      return action;
    }

    const sheetFmt = formatContexts.find((fmt) => fmt.sheetName === actionSheetName(action, context));
    if (!sheetFmt) return action;

    const dateCols = sheetFmt.columns.filter((col) => col.type === 'date' && col.numberFormat);
    if (dateCols.length === 0) return action;
    if (action.format?.numberFormat) return action;

    if (action.type === 'SET_CELL' || action.type === 'SET_FORMULA') {
      const targetCol = action.col;
      if (targetCol === undefined) return action;
      const dateCol = dateCols.find((col) => col.columnIndex === targetCol);
      if (!dateCol?.numberFormat) return action;
      return {
        ...action,
        format: {
          ...action.format,
          numberFormat: dateCol.numberFormat,
        },
      };
    }

    const primaryDateCol = primaryDateColumn(sheetFmt);
    if (!primaryDateCol?.numberFormat) return action;

    return {
      ...action,
      format: {
        ...action.format,
        numberFormat: primaryDateCol.numberFormat,
      },
    };
  });
}

export function summarizeFormatContext(context: WorkbookContext): string {
  return extractFormatContext(context)
    .map((sheetFmt) => {
      const rules = sheetFmt.columns
        .map((col) => `  - Col ${col.columnIndex + 1} "${col.header}": ${col.rule}`)
        .join('\n');
      return `${sheetFmt.sheetName}:\n${rules}`;
    })
    .join('\n\n');
}
