import { buildSheetOverview, formatSheetOverviewMarkdown } from '../src/excel-ai/utils/sheet-overview.util';
import { buildFixture, expectedFacts } from '../eval/usecase-fixture';

/**
 * TASKS.md #223 — "Describe this spreadsheet to me" is answered by this
 * deterministic overview (no LLM), and the live audit caught it stating a date
 * range of 31-12-2023 – 30-11-2024 for data that runs 01-04-2024 – 27-06-2024,
 * and a GST total of ₹3,26,610 against an actual ₹7,09,920. Ground truth comes
 * from the fixture, computed, never hand-typed.
 */
describe('deterministic sheet overview facts (#223)', () => {
  const fixture = buildFixture();
  const facts = expectedFacts();
  const analysis = {
    headers: fixture.sheetData[0] as string[],
    headerRowIndex: 0,
    isEmpty: false,
    rowCount: fixture.sheetData.length,
    columnCount: (fixture.sheetData[0] as string[]).length,
  } as never;
  const overview = buildSheetOverview(fixture.sheetData, analysis, 'Purchase Register');
  const markdown = overview ? formatSheetOverviewMarkdown(overview) : '';

  it('produces an overview at all', () => {
    expect(overview).toBeTruthy();
    expect(markdown).toContain('Purchase Register');
  });

  it('counts the data rows', () => {
    expect(markdown).toContain(String(facts.dataRows));
  });

  it('reports the real date range', () => {
    expect(markdown).toContain(facts.firstDate);
    expect(markdown).toContain(facts.lastDate);
  });

  it('does not invent dates outside the data', () => {
    expect(markdown).not.toMatch(/31-12-2023|30-11-2024/);
  });

  it('reads the date range from the date column, not an ID column', () => {
    expect(overview.dateRange?.column).toBe('Invoice Date');
  });

  it('totals every GST column, not just the first', () => {
    // IGST + CGST + SGST across the fixture.
    expect(markdown).toContain('₹7,09,920');
    expect(markdown).not.toContain('GST: ₹3,26,610');
  });

  it('reports the pre-tax total from the taxable column', () => {
    expect(markdown).toContain('₹42,54,000');
  });
});
