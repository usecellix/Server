import { findHeaderRowIndex } from '../src/excel-ai/utils/header-row-detection.util';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';
import { buildWorkbookContext } from '../src/excel-ai/utils/workbook-context.util';

/**
 * Regression coverage for the header-detection bug: a lone title cell (e.g. a
 * merged "ABC Corp — Purchase Register" row) used to satisfy the bare text-ratio
 * heuristic on its own and get misidentified as the header row, and even a
 * correctly-detected header row was discarded by buildWorkbookContext hardcoding
 * headerRow = 1 regardless of what was found.
 */
describe('header-row-detection.util — findHeaderRowIndex', () => {
  it('does not mistake a single-cell title row for headers', () => {
    const rows = [
      ['ABC Corp — Purchase Register FY24'],
      [],
      ['Date', 'Supplier', 'Invoice No', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 'INV-001', 12500],
    ];
    expect(findHeaderRowIndex(rows, 4)).toBe(2);
  });

  it('skips multiple title/subtitle rows above the real headers', () => {
    const rows = [
      ['Purchase Register'],
      ['For the year ended 31 March 2024'],
      [],
      ['Date', 'Supplier', 'Invoice No', 'Amount', 'Tax %', 'Total'],
      ['01-04-2024', 'Acme Ltd', 'INV-001', 12500, 18, 14750],
    ];
    expect(findHeaderRowIndex(rows, 6)).toBe(3);
  });

  it('still finds headers on row 0 for a normal table', () => {
    const rows = [
      ['Name', 'Age', 'City'],
      ['Alice', 30, 'Mumbai'],
    ];
    expect(findHeaderRowIndex(rows, 3)).toBe(0);
  });

  it('falls back to row 0 when nothing in the scan window qualifies', () => {
    const rows = [[1, 2, 3], [4, 5, 6]];
    expect(findHeaderRowIndex(rows, 3)).toBe(0);
  });
});

describe('SheetAnalyzerService — headerRowIndex end-to-end', () => {
  const analyzer = new SheetAnalyzerService();

  it('detects the header row position, not just the header labels', () => {
    const sheetData = [
      ['Purchase Register — ABC Corp'],
      [],
      ['Date', 'Supplier', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 12500],
      ['02-04-2024', 'Beta Inc', 8500],
    ];
    const analysis = analyzer.analyze(sheetData);
    expect(analysis.headerRowIndex).toBe(2);
    expect(analysis.headers).toEqual(['Date', 'Supplier', 'Amount']);
  });

  it('sumColumn respects headerRowIndex instead of always skipping row 0', () => {
    const sheetData = [
      ['Purchase Register — ABC Corp'],
      [],
      ['Date', 'Supplier', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 12500],
      ['02-04-2024', 'Beta Inc', 8500],
    ];
    const analysis = analyzer.analyze(sheetData);
    const total = analyzer.sumColumn(sheetData, 2, true, analysis.headerRowIndex);
    expect(total).toBe(21000);
  });
});

describe('buildWorkbookContext — no longer hardcodes headerRow = 1', () => {
  const analyzer = new SheetAnalyzerService();

  it('reports the real header row for a sheet with preamble rows', () => {
    const sheetData = [
      ['Purchase Register — ABC Corp'],
      [],
      ['Date', 'Supplier', 'Amount'],
      ['01-04-2024', 'Acme Ltd', 12500],
    ];
    const analysis = analyzer.analyze(sheetData);
    const ctx = buildWorkbookContext(sheetData, analysis);

    // Header is on sheet row 3 (1-based) — this used to always read 1.
    expect(ctx.headerRow).toBe(3);
    expect(ctx.dataStartRow).toBe(4);
    expect(ctx.dataRange.startsWith('A3:')).toBe(true);
  });

  it('still reports row 1 for a normal table with no preamble', () => {
    const sheetData = [
      ['Name', 'Age'],
      ['Alice', 30],
    ];
    const analysis = analyzer.analyze(sheetData);
    const ctx = buildWorkbookContext(sheetData, analysis);

    expect(ctx.headerRow).toBe(1);
    expect(ctx.dataStartRow).toBe(2);
  });
});
