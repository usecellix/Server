import { Injectable } from '@nestjs/common';
import { formatIndianCurrency, formatIndianNumber, parseIndianNumber } from '../utils/indian-format.util';
import { parseFindSearchTerms } from '../utils/find-query-parser.util';
import { SheetAnalysis, SheetAnalyzerService } from './sheet-analyzer.service';
import { WorkbookContext } from '../types/sheet-actions.types';

export interface SelectCellTarget {
  sheetName: string;
  row: number;
  col: number;
}

export interface FindMatch {
  label: string;
  detail?: string;
  sheetName: string;
  /** 0-based row index within the sheet. */
  row: number;
  /** 0-based column index within the sheet. */
  col: number;
  colLetter: string;
  /** 1-based Excel row number. */
  rowNum: number;
  rawValue: string;
}

export interface DataQueryResult {
  answer: string;
  computedValue?: number;
  followUp?: string;
  selectCell?: SelectCellTarget;
  matches?: FindMatch[];
}

export interface FindQueryOptions {
  sheetName?: string;
}

@Injectable()
export class DataQueryService {
  constructor(private readonly analyzer: SheetAnalyzerService) {}

  query(
    subIntent: string | undefined,
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    ctx?: WorkbookContext,
    options?: FindQueryOptions,
  ): DataQueryResult | null {
    const columnIndex = this.analyzer.resolveColumnIndexFromMessage(message, analysis);
    const lower = message.toLowerCase();

    switch (subIntent) {
      case 'find':
        return this.findRows(message, sheetData, analysis, columnIndex, options);
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

  private extractFindSearchTerms(message: string): string[] {
    return parseFindSearchTerms(message);
  }

  private extractFindSearchTerm(message: string): string | null {
    const terms = this.extractFindSearchTerms(message);
    return terms.length ? terms.join(', ') : null;
  }

  /**
   * Numeric find: integer search "148" matches 148, 148.5, 148.00 — not 1487,
   * 148000, or 21486. Decimal search "148.5" matches that value only.
   */
  private numericValueMatchesSearch(absCell: number, absSearch: number, searchText: string): boolean {
    const hasDecimalInSearch = /\.\d+/.test(searchText.trim());

    if (hasDecimalInSearch) {
      return Math.abs(absCell - absSearch) < 0.001;
    }

    const searchInt = Math.floor(absSearch + 1e-9);
    const cellInt = Math.floor(absCell + 1e-9);
    return cellInt === searchInt;
  }

  private cellMatchesSearch(
    cellStr: string,
    searchNum: number | null,
    searchText: string,
  ): boolean {
    const cellNum = parseIndianNumber(cellStr);

    if (searchNum !== null && cellNum !== null) {
      return this.numericValueMatchesSearch(Math.abs(cellNum), Math.abs(searchNum), searchText);
    }

    const normalizedCell = cellStr.toLowerCase().replace(/[\s,₹]/g, '');
    const normalizedSearch = searchText.replace(/[\s,₹]/g, '');

    if (searchNum === null) {
      return normalizedCell.includes(normalizedSearch);
    }

    // Non-numeric cell while searching a number — exact token match only.
    const stripped = normalizedCell.replace(/(dr|cr)$/i, '');
    return stripped === normalizedSearch;
  }

  private scanForMatches(
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    colsToScan: number[],
    searchNum: number | null,
    searchText: string,
  ): Array<{
    rowIndex: number;
    rowNum: number;
    colIndex: number;
    colHeader: string;
    colLetter: string;
    rawValue: string;
  }> {
    const headers = analysis.headers;
    const matches: Array<{
      rowIndex: number;
      rowNum: number;
      colIndex: number;
      colHeader: string;
      colLetter: string;
      rawValue: string;
    }> = [];

    for (let row = 1; row < sheetData.length; row += 1) {
      for (const col of colsToScan) {
        const cell = sheetData[row]?.[col];
        if (cell === undefined || cell === null || cell === '') continue;

        const cellStr = String(cell);
        if (!this.cellMatchesSearch(cellStr, searchNum, searchText)) continue;

        matches.push({
          rowIndex: row,
          rowNum: row + 1,
          colIndex: col,
          colHeader: headers[col] || analysis.columnLetters[col],
          colLetter: analysis.columnLetters[col],
          rawValue: cellStr,
        });
      }
    }

    return matches;
  }

  private buildRowContext(
    rowData: unknown[],
    headers: string[],
    matchCol: number,
    matchValue: string,
    colLetter: string,
    rowNum: number,
  ): string {
    const pick = (patterns: RegExp[]): string | null => {
      for (let i = 0; i < headers.length; i += 1) {
        const header = headers[i] ?? '';
        if (!header || !rowData[i]) continue;
        if (patterns.some((pattern) => pattern.test(header))) {
          return `${header}: ${rowData[i]}`;
        }
      }
      return null;
    };

    const supplier =
      pick([/supplier|vendor|party|name/i]) ??
      headers
        .slice(0, 5)
        .map((h, i) => (rowData[i] ? String(rowData[i]) : null))
        .filter(Boolean)
        .slice(0, 2)
        .join(', ');

    const date = pick([/date/i]);
    const gstin = pick([/gstin/i]);
    const header = headers[matchCol] || 'Value';

    const parts = [
      supplier ? String(supplier).replace(/^[^:]+:\s*/, '') : null,
      gstin ? gstin.replace(/^[^:]+:\s*/, '') : null,
      date ? date.replace(/^[^:]+:\s*/, '') : null,
    ].filter(Boolean);

    const summary = parts.length ? parts.join(', ') : `Row ${rowNum}`;
    return `${summary}\n${colLetter}${rowNum} (${header}): ${matchValue}`;
  }

  /** Public accessor for search terms (single label joins multi-value lists). */
  extractSearchTerm(message: string): string | null {
    return this.extractFindSearchTerm(message);
  }

  extractSearchTerms(message: string): string[] {
    return this.extractFindSearchTerms(message);
  }

  /**
   * Scan a single sheet and return structured matches (used for both
   * single-sheet and cross-sheet workbook search).
   */
  collectMatches(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    sheetName: string,
    hintColumnIndex: number | null = null,
  ): FindMatch[] {
    const searchTerms = this.extractFindSearchTerms(message);
    if (!searchTerms.length) return [];

    const headers = analysis.headers;
    const allCols = Array.from({ length: headers.length }, (_, i) => i);
    const allRowMatches: FindMatch[] = [];

    for (const rawTerm of searchTerms) {
      const searchNum = parseIndianNumber(rawTerm);
      const searchText = rawTerm.toLowerCase();

      let scanned = this.scanForMatches(
        sheetData,
        analysis,
        hintColumnIndex !== null ? [hintColumnIndex] : allCols,
        searchNum,
        searchText,
      );

      if (scanned.length === 0 && hintColumnIndex !== null) {
        scanned = this.scanForMatches(sheetData, analysis, allCols, searchNum, searchText);
      }

      const rowMatches = scanned.map((m) => {
        const rowData = sheetData[m.rowIndex] as unknown[];
        const context = this.buildRowContext(
          rowData,
          headers,
          m.colIndex,
          m.rawValue,
          m.colLetter,
          m.rowNum,
        );
        const [label, ...detailParts] = context.split('\n');
        return {
          label: label || `Row ${m.rowNum}`,
          detail: detailParts.join(' ') || undefined,
          sheetName,
          row: m.rowIndex,
          col: m.colIndex,
          colLetter: m.colLetter,
          rowNum: m.rowNum,
          rawValue: m.rawValue,
        };
      });

      allRowMatches.push(...rowMatches);
    }

    return this.deduplicateMatchesByRow(allRowMatches);
  }

  /** One clickable result per row (merge CGST/SGST hits on the same row). */
  private deduplicateMatchesByRow(matches: FindMatch[]): FindMatch[] {
    const groups = new Map<string, FindMatch[]>();

    for (const match of matches) {
      const key = `${match.sheetName}\0${match.row}`;
      const list = groups.get(key) ?? [];
      list.push(match);
      groups.set(key, list);
    }

    const merged: FindMatch[] = [];

    for (const rowMatches of groups.values()) {
      rowMatches.sort((a, b) => a.col - b.col);
      const primary = rowMatches[0];
      const uniqueValues = [...new Set(rowMatches.map((m) => m.rawValue))];
      const columnLabels = rowMatches
        .map((m) => {
          const headerMatch = m.detail?.match(/\(([A-Za-z@ %^]+)\):/);
          return headerMatch?.[1] ?? m.colLetter;
        })
        .filter(Boolean);

      const valueText = uniqueValues.length === 1 ? uniqueValues[0] : uniqueValues.join(', ');
      const columnsText =
        columnLabels.length > 1
          ? columnLabels.join(', ')
          : columnLabels.length === 1
            ? columnLabels[0]
            : '';

      merged.push({
        ...primary,
        label: primary.label,
        detail: columnsText ? `${valueText} · ${columnsText}` : valueText,
      });
    }

    return merged.sort(
      (a, b) => a.sheetName.localeCompare(b.sheetName) || a.rowNum - b.rowNum,
    );
  }

  /** Build a find result (compact answer + jump targets) from aggregated matches. */
  buildFindResult(terms: string | string[], matches: FindMatch[]): DataQueryResult {
    const termList = Array.isArray(terms) ? terms : [terms];
    const label = termList.join(', ');

    if (matches.length === 0) {
      return {
        answer: `No value matching **${label}** found in the workbook. Tally values often appear as "1,868.41 Dr" — try searching with just the number (e.g. 1868).`,
        matches: [],
      };
    }

    const sheetCount = new Set(matches.map((m) => m.sheetName)).size;
    const acrossSheets = sheetCount > 1 ? ` across ${sheetCount} sheets` : '';
    const rowWord = matches.length === 1 ? 'row' : 'rows';
    const intro =
      matches.length === 1
        ? `Found **${label}** in 1 ${rowWord}`
        : `Found **${label}** in ${matches.length} ${rowWord}${acrossSheets}`;

    return {
      answer: intro,
      matches,
    };
  }

  private findRows(
    message: string,
    sheetData: unknown[][],
    analysis: SheetAnalysis,
    hintColumnIndex: number | null,
    options?: FindQueryOptions,
  ): DataQueryResult | null {
    const terms = this.extractFindSearchTerms(message);
    if (!terms.length) {
      return {
        answer:
          'I could not extract a search value from your message. Please try: "Find CGST value 1868".',
      };
    }

    const sheetName = options?.sheetName ?? 'Sheet1';
    const matches = this.collectMatches(message, sheetData, analysis, sheetName, hintColumnIndex);
    return this.buildFindResult(terms, matches);
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

    const total = this.analyzer.sumColumn(
      sheetData,
      columnIndex,
      true,
      analysis.headerRowIndex ?? 0,
    );
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

    const stats = this.analyzer.columnStats(
      sheetData,
      columnIndex,
      true,
      analysis.headerRowIndex ?? 0,
    );
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

    const stats = this.analyzer.columnStats(
      sheetData,
      columnIndex,
      true,
      analysis.headerRowIndex ?? 0,
    );
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

    const stats = this.analyzer.columnStats(
      sheetData,
      columnIndex,
      true,
      analysis.headerRowIndex ?? 0,
    );
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
      const stats = this.analyzer.columnStats(
      sheetData,
      columnIndex,
      true,
      analysis.headerRowIndex ?? 0,
    );
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
    if (/\b(cgst|sgst|igst|tds|gstin|invoice amount|amount|particulars)\b/.test(lower)) {
      return false;
    }
    return /\b(total|sum|add up)\b/.test(lower) && !/\bcolumn\b/.test(lower);
  }
}
