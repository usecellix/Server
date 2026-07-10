import { DataQueryService } from '../src/excel-ai/services/data-query.service';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';

describe('DataQueryService find term extraction', () => {
  const service = new DataQueryService(new SheetAnalyzerService());

  it('extracts all comma-separated values from a multi-value find prompt', () => {
    expect(service.extractSearchTerms('Find values 2290, 4180, 7515')).toEqual([
      '2290',
      '4180',
      '7515',
    ]);
  });

  it('keeps single-value find behavior for column-specific prompts', () => {
    expect(service.extractSearchTerms('Find CGST value 1868')).toEqual(['1868']);
  });

  it('extracts multiple values joined with and', () => {
    expect(service.extractSearchTerms('Find 2290 and 4180')).toEqual(['2290', '4180']);
  });

  it('extracts text phrases and ignores trailing export clauses', () => {
    expect(
      service.extractSearchTerms(
        'Find all the rows Deva steels and create a new sheet with all that value',
      ),
    ).toEqual(['Deva steels']);
  });

  it('sums CGST column for total questions', () => {
    const analyzer = new SheetAnalyzerService();
    const sheetData = [
      ['Date', 'Particulars', 'CGST', 'SGST'],
      ['2024-01-01', 'Purchase A', '1,868.41 Dr', '1,868.41 Dr'],
      ['2024-01-02', 'Purchase B', '500.00 Dr', '500.00 Dr'],
    ];
    const analysis = analyzer.analyze(sheetData);
    const result = service.query(
      'sum',
      'What is the total CGST in this sheet?',
      sheetData,
      analysis,
    );

    expect(result?.answer).toContain('CGST');
    expect(result?.computedValue).toBeCloseTo(2368.41, 2);
  });

  it('sums CGST when Tally title rows precede headers', () => {
    const analyzer = new SheetAnalyzerService();
    const sheetData = [
      ['ABC Traders Pvt Ltd'],
      ['Purchase Register - FY 2024-25'],
      ['Date', 'Particulars', 'Central Tax (CGST)', 'State Tax (SGST)'],
      ['2024-01-01', 'Purchase A', '1,868.41 Dr', '1,868.41 Dr'],
      ['2024-01-02', 'Purchase B', '500.00 Dr', '500.00 Dr'],
    ];
    const analysis = analyzer.analyze(sheetData, {
      knownHeaders: ['Date', 'Particulars', 'Central Tax (CGST)', 'State Tax (SGST)'],
    });
    const result = service.query(
      'sum',
      'What is the total CGST in this sheet?',
      sheetData,
      analysis,
    );

    expect(result?.answer).toContain('CGST');
    expect(result?.computedValue).toBeCloseTo(2368.41, 2);
  });
});
