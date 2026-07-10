import {
  ColumnPattern,
  FormulaFunctionName,
  ParsedFormula,
  RowPattern,
} from './formula.types';

export function colIndexToLetter(index: number): string {
  let result = '';
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export function detectRowPatterns(
  values: unknown[][],
  parsedFormulas: ParsedFormula[][],
): RowPattern[] {
  return values.map((row, rowIdx) => {
    const rowNum = rowIdx + 1;
    const nonEmpty = row.filter((v) => v !== null && v !== '');
    const formulas = parsedFormulas[rowIdx] ?? [];
    const nonEmptyFormulas = formulas.filter((f) => f.raw.startsWith('='));

    if (nonEmpty.length === 0) {
      return {
        rowIndex: rowNum,
        type: 'empty',
        formulaTypes: [],
        description: `Row ${rowNum}: empty`,
      };
    }

    const stringCount = nonEmpty.filter(
      (v) => typeof v === 'string' && Number.isNaN(Number(v)),
    ).length;
    if (stringCount / nonEmpty.length > 0.6 && nonEmptyFormulas.length === 0) {
      return {
        rowIndex: rowNum,
        type: 'header',
        formulaTypes: [],
        description: `Row ${rowNum}: header row`,
      };
    }

    const aggFormulas = nonEmptyFormulas.filter((f) => f.isAggregation);
    if (
      aggFormulas.length > 0 &&
      aggFormulas.length / Math.max(nonEmptyFormulas.length, 1) > 0.4
    ) {
      const fns = [...new Set(aggFormulas.flatMap((f) => f.functions))];
      return {
        rowIndex: rowNum,
        type: 'total',
        formulaTypes: fns,
        description: `Row ${rowNum}: aggregation row (${fns.join(', ')}) — do not insert rows below without updating these formulas`,
      };
    }

    if (aggFormulas.length > 0) {
      const fns = [...new Set(aggFormulas.flatMap((f) => f.functions))];
      return {
        rowIndex: rowNum,
        type: 'subtotal',
        formulaTypes: fns,
        description: `Row ${rowNum}: subtotal row (${fns.join(', ')})`,
      };
    }

    const formulaTypes = [...new Set(nonEmptyFormulas.flatMap((f) => f.functions))];
    return {
      rowIndex: rowNum,
      type: 'data',
      formulaTypes,
      description: `Row ${rowNum}: data row${formulaTypes.length ? ` with ${formulaTypes.join(', ')}` : ''}`,
    };
  });
}

export function detectColumnPatterns(
  parsedFormulas: ParsedFormula[][],
  columnCount: number,
): ColumnPattern[] {
  const patterns: ColumnPattern[] = [];

  for (let c = 0; c < columnCount; c += 1) {
    const colLetter = colIndexToLetter(c);
    const colFormulas = parsedFormulas
      .map((row) => row[c])
      .filter((f): f is ParsedFormula => Boolean(f?.raw.startsWith('=')));

    if (colFormulas.length === 0) {
      const hasValues = parsedFormulas.some((row) => {
        const cell = row[c];
        return cell?.raw && !cell.raw.startsWith('=');
      });
      if (!hasValues) continue;
      patterns.push({
        colLetter,
        type: 'label',
        description: `Column ${colLetter}: labels/data (no formulas)`,
      });
      continue;
    }

    const allFunctions = colFormulas.flatMap((f) => f.functions);
    const counts: Record<string, number> = {};
    for (const fn of allFunctions) {
      counts[fn] = (counts[fn] ?? 0) + 1;
    }
    const dominant = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
    const totalCells = parsedFormulas.filter((row) => row[c]?.raw).length;
    const dataRatio = colFormulas.length / Math.max(totalCells, 1);

    patterns.push({
      colLetter,
      type: dataRatio > 0.7 ? 'formula' : 'mixed',
      dominantFunction: dominant?.[0] as FormulaFunctionName | undefined,
      description: `Column ${colLetter}: ${dataRatio > 0.7 ? 'formula' : 'mixed'} column${dominant ? ` (mostly ${dominant[0]})` : ''}`,
    });
  }

  return patterns;
}

export function findDataRowRange(
  rowPatterns: RowPattern[],
): { start: number; end: number } | undefined {
  const dataRows = rowPatterns.filter((r) => r.type === 'data').map((r) => r.rowIndex);
  if (dataRows.length === 0) return undefined;
  return { start: Math.min(...dataRows), end: Math.max(...dataRows) };
}
