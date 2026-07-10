import { WorkbookContext as RichWorkbookContext } from '../../types/cellix.types';
import { ChangeSetRecord } from '../../audit/types/change-set.types';

function parseCellKey(key: string): { sheet: string; cell: string } | null {
  const bang = key.indexOf('!');
  if (bang <= 0) return null;
  return {
    sheet: key.slice(0, bang),
    cell: key.slice(bang + 1),
  };
}

function columnLettersToIndex(letters: string): number {
  let index = 0;
  for (const char of letters.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

function parseRowFromCell(cell: string): number {
  const match = /\d+/.exec(cell);
  return match ? Number.parseInt(match[0], 10) - 1 : 0;
}

function parseColFromCell(cell: string): number {
  const match = /^([A-Za-z]+)/.exec(cell);
  return match ? columnLettersToIndex(match[1]) : 0;
}

export interface RefinementContextBundle {
  sheetData: unknown[][];
  richWorkbookContext: RichWorkbookContext;
  promptContext: string;
}

/** Build a sparse workbook snapshot from a prior change set for quick-edit refinements. */
export function buildRefinementContext(changeSet: ChangeSetRecord): RefinementContextBundle {
  const sheetMap = new Map<string, Map<string, unknown>>();

  for (const change of changeSet.changes) {
    const rowKey = `${change.sheet}!${change.cell}`;
    if (!sheetMap.has(change.sheet)) {
      sheetMap.set(change.sheet, new Map());
    }
    sheetMap.get(change.sheet)!.set(change.cell, change.before);
  }

  for (const [key, snapshot] of Object.entries(changeSet.beforeState)) {
    const parsed = parseCellKey(key);
    if (!parsed) continue;
    if (!sheetMap.has(parsed.sheet)) {
      sheetMap.set(parsed.sheet, new Map());
    }
    sheetMap.get(parsed.sheet)!.set(parsed.cell, snapshot.value);
  }

  const sheetNames = Array.from(sheetMap.keys());
  const activeSheet = sheetNames[0] ?? 'Sheet1';
  const activeCells = sheetMap.get(activeSheet) ?? new Map<string, unknown>();

  let maxRow = 0;
  let maxCol = 0;
  for (const cell of activeCells.keys()) {
    maxRow = Math.max(maxRow, parseRowFromCell(cell));
    maxCol = Math.max(maxCol, parseColFromCell(cell));
  }

  const rowCount = Math.max(maxRow + 1, 1);
  const colCount = Math.max(maxCol + 1, 1);
  const values: unknown[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: colCount }, () => ''),
  );

  for (const [cell, value] of activeCells.entries()) {
    const row = parseRowFromCell(cell);
    const col = parseColFromCell(cell);
    if (values[row] && col < values[row].length) {
      values[row][col] = value ?? '';
    }
  }

  const headers = values[0]?.map((cell) => String(cell ?? '')) ?? [];
  const sampleData = values.slice(1, 6).map((row) =>
    row.map((cell) => (cell === '' || cell == null ? null : (cell as string | number))),
  );

  const richWorkbookContext: RichWorkbookContext = {
    activeSheet,
    sheets: sheetNames.map((sheetName) => {
      const cells = sheetMap.get(sheetName) ?? new Map();
      let sheetMaxRow = 0;
      let sheetMaxCol = 0;
      for (const cell of cells.keys()) {
        sheetMaxRow = Math.max(sheetMaxRow, parseRowFromCell(cell));
        sheetMaxCol = Math.max(sheetMaxCol, parseColFromCell(cell));
      }

      return {
        sheetName,
        usedRange: `A1:${String.fromCharCode(65 + Math.min(sheetMaxCol, 25))}${sheetMaxRow + 1}`,
        rowCount: sheetMaxRow + 1,
        colCount: sheetMaxCol + 1,
        headers,
        sampleData,
        columnMeta: [],
      };
    }),
  };

  const changeSummary = changeSet.changes
    .slice(0, 30)
    .map((change) => `${change.sheet}!${change.cell}: ${JSON.stringify(change.before)} → ${JSON.stringify(change.after)}`)
    .join('\n');

  const promptContext = [
    'QUICK EDIT MODE — refining a prior audited change set without a full workbook read.',
    `Prior change set: ${changeSet.changeSetId}`,
    `Original prompt: ${changeSet.prompt}`,
    `Status: ${changeSet.status}`,
    'Affected cells (before → after):',
    changeSummary,
    'Apply only the user adjustment on top of this prior change context.',
  ].join('\n');

  return {
    sheetData: values,
    richWorkbookContext,
    promptContext,
  };
}
