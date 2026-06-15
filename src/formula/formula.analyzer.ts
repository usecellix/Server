import { Injectable } from '@nestjs/common';
import { SheetContext } from '../agents/types/agent.types';
import { buildDependencyGraph } from './dependency.graph';
import { parseFormula } from './formula.parser';
import {
  detectColumnPatterns,
  detectRowPatterns,
  findDataRowRange,
  colIndexToLetter,
} from './pattern.detector';
import { FormulaInsights, FormulaNode, RowPattern } from './formula.types';

@Injectable()
export class FormulaAnalyzer {
  analyzeSheet(sheet: SheetContext): FormulaInsights {
    const parsedFormulas = sheet.formulas.map((row) =>
      row.map((cell) => parseFormula(typeof cell === 'string' ? cell : '')),
    );

    const formulaNodes = this.buildFormulaNodes(sheet.name, parsedFormulas);
    buildDependencyGraph(formulaNodes);

    const allParsed = parsedFormulas.flat().filter((f) => f.raw.startsWith('='));
    const totalFormulas = allParsed.length;

    const crossSheetRefs = [...new Set(allParsed.flatMap((f) => f.crossSheetRefs))];

    const functionsSummary: Record<string, number> = {};
    for (const f of allParsed) {
      for (const fn of f.functions) {
        functionsSummary[fn] = (functionsSummary[fn] ?? 0) + 1;
      }
    }

    const rowPatterns = detectRowPatterns(sheet.values, parsedFormulas);
    const columnPatterns = detectColumnPatterns(parsedFormulas, sheet.columnCount);
    const aggregationRows = rowPatterns
      .filter((r) => r.type === 'total' || r.type === 'subtotal')
      .map((r) => r.rowIndex);
    const dataRowRange = findDataRowRange(rowPatterns);

    const dependencyWarnings = this.buildWarnings(rowPatterns, aggregationRows, dataRowRange);

    const llmSummary = this.buildLLMSummary(
      sheet.name,
      totalFormulas,
      crossSheetRefs,
      functionsSummary,
      rowPatterns,
      columnPatterns,
      aggregationRows,
      dataRowRange,
      dependencyWarnings,
    );

    return {
      sheetName: sheet.name,
      totalFormulas,
      crossSheetRefs,
      functionsSummary,
      rowPatterns,
      columnPatterns,
      aggregationRows,
      dataRowRange,
      dependencyWarnings,
      llmSummary,
    };
  }

  private buildFormulaNodes(
    sheetName: string,
    parsedFormulas: ReturnType<typeof parseFormula>[][],
  ): FormulaNode[] {
    const nodes: FormulaNode[] = [];
    for (let r = 0; r < parsedFormulas.length; r += 1) {
      for (let c = 0; c < parsedFormulas[r].length; c += 1) {
        const formula = parsedFormulas[r][c];
        if (!formula.raw.startsWith('=')) continue;
        nodes.push({
          address: `${colIndexToLetter(c)}${r + 1}`,
          sheetName,
          formula,
          rowIndex: r + 1,
          colIndex: c,
        });
      }
    }
    return nodes;
  }

  private buildWarnings(
    rowPatterns: RowPattern[],
    aggregationRows: number[],
    dataRowRange?: { start: number; end: number },
  ): string[] {
    const warnings: string[] = [];

    if (aggregationRows.length > 0 && dataRowRange) {
      const dangerRows = aggregationRows.filter((r) => r === dataRowRange.end + 1);
      if (dangerRows.length > 0) {
        warnings.push(
          `Row ${dangerRows[0]} is an aggregation row immediately after data. ` +
            `Inserting a new data row at row ${dataRowRange.end + 1} will push this total down — ` +
            `insert at row ${dataRowRange.end} or earlier instead.`,
        );
      }
    }

    const totalRows = rowPatterns.filter((r) => r.type === 'total');
    if (totalRows.length > 0) {
      warnings.push(
        `Total rows at: ${totalRows.map((r) => r.rowIndex).join(', ')}. ` +
          `Do not insert rows below the last data row without updating these aggregation formulas.`,
      );
    }

    return warnings;
  }

  private buildLLMSummary(
    sheetName: string,
    totalFormulas: number,
    crossSheetRefs: string[],
    functionsSummary: Record<string, number>,
    rowPatterns: RowPattern[],
    columnPatterns: ReturnType<typeof detectColumnPatterns>,
    aggregationRows: number[],
    dataRowRange?: { start: number; end: number },
    warnings: string[] = [],
  ): string {
    if (totalFormulas === 0) {
      return `=== Formula Analysis: ${sheetName} ===\nNo formulas on this sheet.`;
    }

    const lines: string[] = [];
    lines.push(`=== Formula Analysis: ${sheetName} ===`);
    lines.push(`Total formulas: ${totalFormulas}`);

    if (Object.keys(functionsSummary).length) {
      const topFns = Object.entries(functionsSummary)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 6)
        .map(([fn, count]) => `${fn}(${count})`);
      lines.push(`Functions used: ${topFns.join(', ')}`);
    }

    if (crossSheetRefs.length) {
      lines.push(`Cross-sheet refs: ${crossSheetRefs.join(', ')}`);
    }

    if (dataRowRange) {
      lines.push(`Data rows: ${dataRowRange.start}–${dataRowRange.end}`);
    }

    if (aggregationRows.length) {
      lines.push(`Aggregation/total rows: ${aggregationRows.join(', ')}`);
    }

    const headerRows = rowPatterns.filter((r) => r.type === 'header').map((r) => r.rowIndex);
    if (headerRows.length) {
      lines.push(`Header rows: ${headerRows.join(', ')}`);
    }

    const formulaCols = columnPatterns.filter((c) => c.type === 'formula');
    if (formulaCols.length) {
      lines.push(
        `Formula columns: ${formulaCols
          .map((c) => `${c.colLetter}${c.dominantFunction ? `(${c.dominantFunction})` : ''}`)
          .join(', ')}`,
      );
    }

    if (warnings.length) {
      lines.push('WARNINGS:');
      for (const w of warnings) {
        lines.push(`  ⚠ ${w}`);
      }
    }

    return lines.join('\n');
  }
}
