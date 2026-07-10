import { Injectable } from '@nestjs/common';
import { SheetActionPayload } from '../types/sheet-actions.types';
import { suggestExportSheetName } from '../utils/find-query-parser.util';
import { DataQueryService, FindMatch } from './data-query.service';
import { SheetAnalysis } from './sheet-analyzer.service';

export interface FindExportSheetSlice {
  sheetName: string;
  sheetData: unknown[][];
  analysis: SheetAnalysis;
  matches: FindMatch[];
}

export interface FindExportPlan {
  answer: string;
  explanation: string;
  actions: SheetActionPayload[];
  matchCount: number;
  exportSheetName: string;
}

@Injectable()
export class FindExportService {
  constructor(private readonly dataQuery: DataQueryService) {}

  buildPlan(message: string, slices: FindExportSheetSlice[]): FindExportPlan | null {
    const terms = this.dataQuery.extractSearchTerms(message);
    if (!terms.length) return null;

    const label = terms.join(', ');
    const matchedRows = this.collectMatchedRows(slices);
    if (matchedRows.length === 0) {
      return {
        answer: `No rows matching **${label}** were found in the workbook.`,
        explanation: 'No matching rows to export.',
        actions: [],
        matchCount: 0,
        exportSheetName: suggestExportSheetName(message, label),
      };
    }

    const multiSource = new Set(matchedRows.map((row) => row.sourceSheet)).size > 1;
    const headers = this.buildHeaders(slices[0]?.analysis, multiSource);
    const rows = matchedRows.map((entry) => {
      const values = this.normalizeRowValues(entry.rowData, headers.length - (multiSource ? 1 : 0));
      return multiSource ? [entry.sourceSheet, ...values] : values;
    });

    const exportSheetName = suggestExportSheetName(message, label);
    const rowWord = matchedRows.length === 1 ? 'row' : 'rows';

    return {
      answer: `Found **${label}** in ${matchedRows.length} ${rowWord} and prepared sheet **${exportSheetName}** with the full row data.`,
      explanation: `Create "${exportSheetName}" and copy ${matchedRows.length} matching ${rowWord} from the workbook.`,
      matchCount: matchedRows.length,
      exportSheetName,
      actions: [
        {
          type: 'CREATE_SHEET',
          sheetName: exportSheetName,
          relativeTo: slices[0]?.sheetName,
          position: 'after',
        },
        {
          type: 'WRITE_TABLE',
          sheetName: exportSheetName,
          headers,
          rows,
        },
      ],
    };
  }

  private collectMatchedRows(
    slices: FindExportSheetSlice[],
  ): Array<{ sourceSheet: string; rowIndex: number; rowData: unknown[] }> {
    const seen = new Set<string>();
    const collected: Array<{ sourceSheet: string; rowIndex: number; rowData: unknown[] }> = [];

    for (const slice of slices) {
      const uniqueRows = [...new Set(slice.matches.map((match) => match.row))].sort((a, b) => a - b);
      for (const rowIndex of uniqueRows) {
        const key = `${slice.sheetName}\0${rowIndex}`;
        if (seen.has(key)) continue;
        seen.add(key);
        const rowData = slice.sheetData[rowIndex];
        if (!rowData) continue;
        collected.push({
          sourceSheet: slice.sheetName,
          rowIndex,
          rowData: rowData as unknown[],
        });
      }
    }

    return collected.sort(
      (a, b) =>
        a.sourceSheet.localeCompare(b.sourceSheet) || a.rowIndex - b.rowIndex,
    );
  }

  private buildHeaders(analysis: SheetAnalysis | undefined, includeSourceSheet: boolean): string[] {
    const base =
      analysis?.headers?.filter((header) => String(header ?? '').trim().length > 0) ??
      analysis?.columnLetters ??
      [];
    const headers = base.length > 0 ? base.map((header) => String(header)) : ['Column A'];
    return includeSourceSheet ? ['Source Sheet', ...headers] : headers;
  }

  private normalizeRowValues(rowData: unknown[], columnCount: number): unknown[] {
    const width = Math.max(columnCount, rowData.length, 1);
    return Array.from({ length: width }, (_, index) => {
      const value = rowData[index];
      return value === undefined || value === null ? '' : value;
    });
  }
}
