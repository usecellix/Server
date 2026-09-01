import { SheetContext, WorkbookContext } from '../agents/types/agent.types';
import { ConditionalFormatRuleInfo } from '../types/cellix.types';
import { ShadowCell, ShadowSheet, ShadowWorkbook } from './shadowWorkbook.types';

export function colIndexToLetter(index: number): string {
  let result = '';
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export function letterToColIndex(letter: string): number {
  return letter
    .toUpperCase()
    .split('')
    .reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
}

export function buildShadowWorkbook(context: WorkbookContext): ShadowWorkbook {
  const sheets = new Map<string, ShadowSheet>();
  for (const sheet of context.sheets) {
    sheets.set(sheet.name, buildShadowSheet(sheet));
  }

  const namedRanges = new Map<string, string>(
    context.namedRanges.map((n) => [n.name, n.formula]),
  );

  return {
    activeSheetName: context.activeSheetName,
    sheets,
    namedRanges,
    tables: [...context.tables],
    changedCells: new Set(),
    structuralLog: [],
    conditionalFormats: [],
  };
}

function buildShadowSheet(sheet: SheetContext): ShadowSheet {
  const cells = new Map<string, ShadowCell>();

  for (let r = 0; r < sheet.values.length; r += 1) {
    const row = sheet.values[r] ?? [];
    for (let c = 0; c < row.length; c += 1) {
      const address = `${colIndexToLetter(c)}${r + 1}`;
      const format = sheet.formats?.[r]?.[c];
      cells.set(address, {
        value: row[c],
        formula: sheet.formulas?.[r]?.[c] ?? '',
        numberFormat: sheet.numberFormats?.[r]?.[c] ?? 'General',
        bold: format?.bold,
        italic: format?.italic,
        fontColor: format?.fontColor,
        fillColor: format?.fillColor,
      });
    }
  }

  return {
    name: sheet.name,
    cells,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    structure: sheet.structure,
  };
}

export function shadowToWorkbookContext(shadow: ShadowWorkbook): WorkbookContext {
  const sheets = Array.from(shadow.sheets.values()).map(shadowSheetToContext);
  return {
    activeSheetName: shadow.activeSheetName,
    sheets,
    namedRanges: Array.from(shadow.namedRanges.entries()).map(([name, formula]) => ({
      name,
      formula,
    })),
    tables: [...shadow.tables],
    // The shadow only tracks presence (sheet/range/kind, #39) — no real id or
    // summary exists mid-turn until a live Office.js read happens (#38), so
    // `id` is deliberately left unset rather than fabricated as targetable.
    conditionalFormats: shadow.conditionalFormats.map((cf) => ({
      id: '',
      sheetName: cf.sheetName,
      range: cf.range,
      ruleKind: cf.ruleKind as ConditionalFormatRuleInfo['ruleKind'],
      summary: cf.ruleKind,
    })),
  };
}

export function shadowSheetToContext(sheet: ShadowSheet): SheetContext {
  const addresses = Array.from(sheet.cells.keys());
  let maxRow = sheet.rowCount;
  let maxCol = sheet.columnCount;

  if (addresses.length > 0) {
    maxRow = Math.max(
      maxRow,
      ...addresses.map((a) => Number.parseInt(a.replace(/[A-Z]+/i, ''), 10)),
    );
    maxCol = Math.max(
      maxCol,
      ...addresses.map((a) => letterToColIndex(a.replace(/\d+/g, '')) + 1),
    );
  }

  const values: unknown[][] = Array.from({ length: maxRow }, () =>
    Array.from({ length: maxCol }, () => null),
  );
  const formulas: string[][] = Array.from({ length: maxRow }, () =>
    Array.from({ length: maxCol }, () => ''),
  );
  const numberFormats: string[][] = Array.from({ length: maxRow }, () =>
    Array.from({ length: maxCol }, () => 'General'),
  );
  const formats: NonNullable<SheetContext['formats']> = Array.from({ length: maxRow }, () =>
    Array.from({ length: maxCol }, () => ({})),
  );

  for (const [addr, cell] of sheet.cells) {
    const col = letterToColIndex(addr.replace(/\d+/g, ''));
    const row = Number.parseInt(addr.replace(/[A-Z]+/i, ''), 10) - 1;
    if (row >= 0 && row < maxRow && col >= 0 && col < maxCol) {
      values[row][col] = cell.formula ? cell.formula : cell.value;
      formulas[row][col] = cell.formula;
      numberFormats[row][col] = cell.numberFormat;
      formats[row][col] = {
        bold: cell.bold,
        italic: cell.italic,
        fontColor: cell.fontColor,
        fillColor: cell.fillColor,
      };
    }
  }

  return {
    name: sheet.name,
    usedRange: maxRow > 0 && maxCol > 0 ? `A1:${colIndexToLetter(maxCol - 1)}${maxRow}` : '',
    rowCount: maxRow,
    columnCount: maxCol,
    values,
    formulas,
    numberFormats,
    formats,
    structure: sheet.structure as SheetContext['structure'],
    // Shadow workbook diffing doesn't reason about title rows — this is an
    // internal post-apply verification snapshot, not user-facing header context.
    headerRowIndex: 0,
  };
}

export function deepCloneShadow(shadow: ShadowWorkbook): ShadowWorkbook {
  const sheets = new Map<string, ShadowSheet>();
  for (const [name, sheet] of shadow.sheets) {
    sheets.set(name, {
      ...sheet,
      cells: new Map(sheet.cells),
    });
  }
  return {
    activeSheetName: shadow.activeSheetName,
    sheets,
    namedRanges: new Map(shadow.namedRanges),
    tables: [...shadow.tables],
    changedCells: new Set(shadow.changedCells),
    structuralLog: [...shadow.structuralLog],
    conditionalFormats: shadow.conditionalFormats.map((cf) => ({ ...cf })),
  };
}
