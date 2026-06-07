import { Injectable } from '@nestjs/common';
import { formatIndianCurrency, formatIndianNumber, parseIndianNumber } from '../utils/indian-format.util';
import { SheetAnalysis, SheetAnalyzerService } from './sheet-analyzer.service';
import { WorkbookContext } from '../types/sheet-actions.types';

export interface DataQueryResult {
  answer: string;
  computedValue?: number;
  followUp?: string;
}

@Injectable()
export class DataQueryService {
  constructor(private readonly analyzer: SheetAnalyzerService) {}

  query(
    subIntent: string | undefined,
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    ctx: WorkbookContext,
  ): DataQueryResult | null {
    const columnIndex = this.analyzer.resolveColumnIndexFromMessage(message, analysis);
    const lower = message.toLowerCase();

    switch (subIntent) {
      case 'average':
        return this.average(columnIndex, sheetData, analysis);
      case 'max':
        return this.max(columnIndex, sheetData, analysis);
      case 'min':
        return this.min(columnIndex, sheetData, analysis);
      case 'count':
        return this.count(columnIndex, sheetData, analysis, lower);
      case 'blank':
        return this.countBlank(columnIndex, sheetData, analysis);
      case 'duplicate':
        return this.findDuplicates(columnIndex, sheetData, analysis);
      case 'percentage':
        return this.percentage(columnIndex, sheetData, analysis, lower);
      case 'sum':
      default:
        return this.sum(columnIndex, sheetData, analysis, lower);
    }
  }

  private sum(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    lower: string,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      if (this.isGenericTotal(lower)) {
        return {
          answer:
            'Which column should I total? Tell me the column name or letter (e.g. "Invoice Amount" or column D).',
        };
      }
      return null;
    }

    const total = this.analyzer.sumColumn(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];
    const letter = analysis.columnLetters[columnIndex];

    return {
      answer: `The total for **${label}** (column ${letter}) is ${formatIndianCurrency(total)}.`,
      computedValue: total,
      followUp: `Want me to add a SUM formula at the bottom of column ${letter}?`,
    };
  }

  private average(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I average? Tell me the column name or letter.' };
    }

    const stats = this.analyzer.columnStats(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    return {
      answer: `The average for **${label}** is ${formatIndianCurrency(stats.average)} (${stats.count} numeric values).`,
      computedValue: stats.average,
      followUp: `Want me to add an AVERAGE formula for column ${analysis.columnLetters[columnIndex]}?`,
    };
  }

  private max(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I check for the maximum value?' };
    }

    const stats = this.analyzer.columnStats(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    if (stats.max === null) {
      return { answer: `Column **${label}** has no numeric values.` };
    }

    const rowInfo = stats.maxRow !== null ? ` (row ${stats.maxRow + 1})` : '';
    return {
      answer: `The highest value in **${label}** is ${formatIndianCurrency(stats.max)}${rowInfo}.`,
      computedValue: stats.max,
      followUp: 'Want me to highlight the row with the highest value?',
    };
  }

  private min(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I check for the minimum value?' };
    }

    const stats = this.analyzer.columnStats(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    if (stats.min === null) {
      return { answer: `Column **${label}** has no numeric values.` };
    }

    const rowInfo = stats.minRow !== null ? ` (row ${stats.minRow + 1})` : '';
    return {
      answer: `The lowest value in **${label}** is ${formatIndianCurrency(stats.min)}${rowInfo}.`,
      computedValue: stats.min,
    };
  }

  private count(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    lower: string,
  ): DataQueryResult | null {
    if (columnIndex !== null) {
      const stats = this.analyzer.columnStats(sheetData, columnIndex);
      const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];
      const threshold = this.extractThreshold(lower);

      if (threshold !== null) {
        const above = this.countAboveThreshold(sheetData, columnIndex, threshold);
        return {
          answer: `${formatIndianNumber(above)} rows in **${label}** are above ${formatIndianCurrency(threshold)}.`,
          computedValue: above,
        };
      }

      return {
        answer: `Column **${label}** has ${formatIndianNumber(stats.nonEmpty)} non-empty cells (${formatIndianNumber(stats.count)} numeric).`,
        computedValue: stats.nonEmpty,
      };
    }

    const dataRows = Math.max(0, analysis.rowCount - 1);
    return {
      answer: `There are **${formatIndianNumber(dataRows)}** data rows (${formatIndianNumber(analysis.rowCount)} total including header).`,
      computedValue: dataRows,
    };
  }

  private countBlank(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I check for blank cells?' };
    }

    const blank = this.analyzer.countBlank(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    return {
      answer: `Column **${label}** has ${formatIndianNumber(blank)} blank cell(s).`,
      computedValue: blank,
      followUp: blank > 0 ? 'Want me to delete all rows where this column is blank?' : undefined,
    };
  }

  private findDuplicates(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I check for duplicates?' };
    }

    const dupes = this.analyzer.findDuplicates(sheetData, columnIndex);
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    if (dupes.length === 0) {
      return { answer: `No duplicate values found in **${label}**.` };
    }

    const preview = dupes
      .slice(0, 5)
      .map((d) => `"${d.value}" (${d.count} times)`)
      .join(', ');
    const more = dupes.length > 5 ? ` and ${dupes.length - 5} more` : '';

    return {
      answer: `Found ${formatIndianNumber(dupes.length)} duplicate value(s) in **${label}**: ${preview}${more}.`,
      followUp: 'Want me to highlight duplicate values?',
    };
  }

  private percentage(
    columnIndex: number | null,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    lower: string,
  ): DataQueryResult | null {
    if (columnIndex === null) {
      return { answer: 'Which column should I calculate the percentage for?' };
    }

    const threshold = this.extractThreshold(lower) ?? 100000;
    const total = this.analyzer.countDataRows(sheetData);
    const above = this.countAboveThreshold(sheetData, columnIndex, threshold);
    const pct = total > 0 ? (above / total) * 100 : 0;
    const label = analysis.headers[columnIndex] || analysis.columnLetters[columnIndex];

    return {
      answer: `${formatIndianNumber(pct, 1)}% of rows (${formatIndianNumber(above)} of ${formatIndianNumber(total)}) in **${label}** are above ${formatIndianCurrency(threshold)}.`,
      computedValue: pct,
    };
  }

  private countAboveThreshold(sheetData: unknown[][], columnIndex: number, threshold: number): number {
    let count = 0;
    for (let row = 1; row < sheetData.length; row += 1) {
      const num = parseIndianNumber(sheetData[row]?.[columnIndex]);
      if (num !== null && num > threshold) count += 1;
    }
    return count;
  }

  private extractThreshold(lower: string): number | null {
    const lakhMatch = /(\d+)\s*lakh/i.exec(lower);
    if (lakhMatch) return Number(lakhMatch[1]) * 100000;

    const kMatch = /(\d+)\s*k\b/i.exec(lower);
    if (kMatch) return Number(kMatch[1]) * 1000;

    const numMatch = /(?:above|over|greater than|>)\s*[\u20B9]?([\d,]+)/i.exec(lower);
    if (numMatch) {
      const parsed = parseIndianNumber(numMatch[1]);
      if (parsed !== null) return parsed;
    }

    return null;
  }

  private isGenericTotal(lower: string): boolean {
    return /\b(total|sum|add up)\b/.test(lower) && !/\bcolumn\b/.test(lower);
  }
}
