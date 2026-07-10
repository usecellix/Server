import { SheetAnalysis } from '../../excel-ai/services/sheet-analyzer.service';
import { WorkbookContext as RichWorkbookContext } from '../../types/cellix.types';
import { SheetContext, WorkbookContext } from '../types/agent.types';

function inferStructure(
  headers: string[],
  structureHint?: SheetContext['structure'],
): SheetContext['structure'] {
  if (structureHint && structureHint !== 'unknown') {
    return structureHint;
  }
  const lower = headers.map((h) => h.toLowerCase());
  if (
    lower.some((h) => h.includes('gst') || h.includes('revenue') || h.includes('profit')) &&
    lower.some((h) => h.includes('amount') || h.includes('total'))
  ) {
    return 'financial_model';
  }
  if (headers.length >= 3 && headers.every((h) => h.length > 0)) {
    return 'data_table';
  }
  if (lower.some((h) => h.includes('report') || h.includes('summary'))) {
    return 'report';
  }
  return 'unknown';
}

function sheetDataHasContent(sheetData: unknown[][]): boolean {
  return sheetData.some(
    (row) =>
      Array.isArray(row) &&
      row.some((cell) => cell !== null && cell !== '' && String(cell).trim() !== ''),
  );
}

function buildSheetContext(
  name: string,
  sheetData: unknown[][],
  analysis: SheetAnalysis,
  richSheet?: RichWorkbookContext['sheets'][number],
): SheetContext {
  const values = sheetData.map((row) => (Array.isArray(row) ? [...row] : []));
  const declaredRowCount = richSheet?.rowCount ?? analysis.rowCount;
  const columnCount = Math.max(
    richSheet?.colCount ?? analysis.columnCount,
    analysis.columnCount,
    values.reduce((max, row) => Math.max(max, row.length), 0),
  );
  const formulas: string[][] = values.map((row) =>
    Array.from({ length: columnCount }, (_, col) => {
      const cell = row[col];
      return typeof cell === 'string' && cell.startsWith('=') ? cell : '';
    }),
  );
  const numberFormats: string[][] = Array.from({ length: values.length }, () =>
    Array.from({ length: columnCount }, (_, colIdx) => {
      const meta = richSheet?.columnMeta?.[colIdx];
      return meta?.numberFormat ?? '';
    }),
  );

  const usedRange =
    richSheet?.usedRange ??
    (analysis.isEmpty
      ? 'A1'
      : `A1:${analysis.columnLetters[analysis.columnCount - 1] ?? 'A'}${declaredRowCount}`);

  const dataTruncated = Boolean(
    richSheet?.compressionMeta?.truncated ||
      declaredRowCount > values.length,
  );

  return {
    name,
    usedRange,
    rowCount: Math.max(declaredRowCount, values.length),
    columnCount,
    values,
    formulas,
    numberFormats,
    structure: inferStructure(analysis.headers, richSheet?.structure),
    compressionMeta: richSheet?.compressionMeta,
    dataTruncated,
  };
}

export function buildAgentWorkbookContext(
  richContext: RichWorkbookContext,
  sheetData: unknown[][],
  analysis: SheetAnalysis,
): WorkbookContext {
  const activeSheetName = richContext.activeSheet;
  const richSheets = richContext.sheets;
  const onDemandFetchEnabled = richSheets.some(
    (sheet) => sheet.compressionMeta?.onDemandFetchEnabled,
  );

  const sheets: SheetContext[] = richSheets.map((richSheet) => {
    const isActive = richSheet.sheetName === activeSheetName;
    const data = isActive
      ? sheetDataHasContent(sheetData)
        ? sheetData
        : reconstructFromSnapshot(richSheet)
      : reconstructFromSnapshot(richSheet);
    const sheetAnalysis = isActive
      ? {
          ...analysis,
          rowCount: Math.max(analysis.rowCount, richSheet.rowCount),
          columnCount: Math.max(analysis.columnCount, richSheet.colCount),
        }
      : {
          rowCount: richSheet.rowCount,
          columnCount: richSheet.colCount,
          headers: richSheet.headers,
          headerRowIndex: 0,
          isEmpty: richSheet.rowCount === 0,
          columnLetters: Array.from({ length: richSheet.colCount }, (_, i) =>
            String.fromCharCode(65 + (i % 26)),
          ),
        };
    return buildSheetContext(richSheet.sheetName, data, sheetAnalysis, richSheet);
  });

  return {
    activeSheetName,
    sheets,
    namedRanges: (richContext.namedRanges ?? []).map((n) => ({
      name: n.name,
      formula: n.formula,
    })),
    tables: (richContext.tables ?? []).map((t) => t.name),
    selectedRange: richContext.selectedRange,
    onDemandFetchEnabled,
  };
}

function reconstructFromSnapshot(snapshot: RichWorkbookContext['sheets'][number]): unknown[][] {
  const rows: unknown[][] = [];
  if (snapshot.headers.length > 0) {
    rows.push([...snapshot.headers]);
  }
  for (const sampleRow of snapshot.sampleData ?? []) {
    rows.push([...sampleRow]);
  }
  return rows;
}
