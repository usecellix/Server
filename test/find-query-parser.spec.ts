import {
  isDataAggregationMessage,
  parseFindSearchTerms,
  resolveLocalFindRoute,
  suggestExportSheetName,
  stripFindFollowOnClauses,
} from '../src/excel-ai/utils/find-query-parser.util';
import { DataQueryService } from '../src/excel-ai/services/data-query.service';
import { FindExportService } from '../src/excel-ai/services/find-export.service';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';

describe('find-query-parser', () => {
  it('routes pure find prompts to read_only', () => {
    expect(resolveLocalFindRoute('Find CGST value 1868')).toBe('read_only');
    expect(resolveLocalFindRoute('Where is invoice 2290')).toBe('read_only');
  });

  it('routes find + export prompts to export_rows', () => {
    const prompt =
      'Find all the rows Deva steels and create a new sheet with all that value';
    expect(resolveLocalFindRoute(prompt)).toBe('export_rows');
    expect(
      resolveLocalFindRoute('find Applied and copy those rows to a new sheet'),
    ).toBe('export_rows');
  });

  it('parses Applied find+copy search term without the export clause', () => {
    expect(
      parseFindSearchTerms('find Applied and copy those rows to a new sheet'),
    ).toEqual(['Applied']);
  });

  it('strips trailing export clauses before parsing search terms', () => {
    const prompt =
      'Find all the rows Deva steels and create a new sheet with all that value';
    expect(stripFindFollowOnClauses(prompt)).toBe('Find all the rows Deva steels');
    expect(parseFindSearchTerms(prompt)).toEqual(['Deva steels']);
  });

  it('extracts email addresses as find search terms', () => {
    expect(parseFindSearchTerms('lookup the value techpath786@gmail.com')).toEqual([
      'techpath786@gmail.com',
    ]);
    expect(parseFindSearchTerms('find anita.george@example.com')).toEqual([
      'anita.george@example.com',
    ]);
  });

  it('strips leading "value" noise from text find phrases', () => {
    expect(parseFindSearchTerms('lookup the value qqqq')).toEqual(['qqqq']);
  });

  it('does not treat aggregation questions as find search terms', () => {
    expect(isDataAggregationMessage('What is the total CGST in this sheet?')).toBe(true);
    expect(parseFindSearchTerms('What is the total CGST in this sheet?')).toEqual([]);
  });

  it('suggests a sheet name from the search label when none is provided', () => {
    expect(
      suggestExportSheetName(
        'Find "Deva steels" and create a new sheet with those rows',
        'Deva steels',
      ),
    ).toBe('Deva steels');
  });
});

describe('FindExportService', () => {
  const service = new FindExportService(new DataQueryService(new SheetAnalyzerService()));

  it('builds CREATE_SHEET + WRITE_TABLE actions for matched rows', () => {
    const sheetData = [
      ['Supplier', 'Amount'],
      ['Deva steels', 100],
      ['Other vendor', 200],
      ['Deva steels', 300],
    ];
    const analysis = new SheetAnalyzerService().analyze(sheetData);
    const matches = new DataQueryService(new SheetAnalyzerService()).collectMatches(
      'Find all the rows Deva steels and create a new sheet',
      sheetData,
      analysis,
      'Purchase register',
    );

    const plan = service.buildPlan(
      'Find all the rows Deva steels and create a new sheet',
      [
        {
          sheetName: 'Purchase register',
          sheetData,
          analysis,
          matches,
        },
      ],
    );

    expect(plan?.matchCount).toBe(2);
    expect(plan?.actions).toEqual([
      {
        type: 'CREATE_SHEET',
        sheetName: 'Deva steels',
        relativeTo: 'Purchase register',
        position: 'after',
      },
      {
        type: 'WRITE_TABLE',
        sheetName: 'Deva steels',
        headers: ['Supplier', 'Amount'],
        rows: [
          ['Deva steels', 100],
          ['Deva steels', 300],
        ],
      },
    ]);
  });
});
