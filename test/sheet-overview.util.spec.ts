import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';
import {
  buildSheetOverview,
  formatSheetOverviewMarkdown,
  isSheetOverviewRequest,
  sanitizeAskAnswer,
} from '../src/excel-ai/utils/sheet-overview.util';

/**
 * Purchase Register–style fixture with known hand-calculable aggregates:
 * Qty = 2+1+3 = 6
 * Tax Amount = 18+9+36 = 63
 * Total Amount = 118+59+236 = 413
 * Payment Status: Paid (2 rows, 118+236=354), Pending (1 row, 59)
 * Largest supplier: Sigma (₹236)
 */
const PURCHASE_REGISTER: unknown[][] = [
  [
    'Date',
    'Supplier',
    'Qty',
    'Unit Price',
    'Tax Amount',
    'Total Amount',
    'Payment Status',
    'Remarks',
  ],
  ['01-04-2024', 'Acme Traders', 2, 50, 18, 118, 'Paid', ''],
  ['15-04-2024', 'Beta Corp', 1, 50, 9, 59, 'Pending', ''],
  ['30-05-2024', 'Sigma Ltd', 3, 200 / 3, 36, 236, 'Paid', ''],
];

describe('sheet-overview.util (Spec 23)', () => {
  const analyzer = new SheetAnalyzerService();

  function overviewFor(rows: unknown[][]) {
    const analysis = analyzer.analyze(rows);
    return buildSheetOverview(rows, analysis, 'Purchase Register');
  }

  describe('isSheetOverviewRequest', () => {
    it.each([
      'Tell me about this sheet',
      'summarize this data',
      "what's in this sheet?",
      'Give me an overview of the sheet',
      'describe this spreadsheet',
    ])('matches broad phrasing: %s', (msg) => {
      expect(isSheetOverviewRequest(msg)).toBe(true);
    });

    it.each(['what does Qty mean', 'sum of total amount', 'find Acme', 'sort by date'])(
      'rejects narrower asks: %s',
      (msg) => {
        expect(isSheetOverviewRequest(msg)).toBe(false);
      },
    );
  });

  describe('buildSheetOverview aggregates', () => {
    it('computes row count, date range, numeric totals, status breakdown, top supplier', () => {
      const overview = overviewFor(PURCHASE_REGISTER);

      expect(overview.rowCount).toBe(3);
      expect(overview.sheetName).toBe('Purchase Register');
      expect(overview.dateRange).toEqual({
        column: 'Date',
        from: '01-04-2024',
        to: '30-05-2024',
      });

      const qty = overview.numericSummaries.find((s) => s.column === 'Qty');
      const tax = overview.numericSummaries.find((s) => s.column === 'Tax Amount');
      const total = overview.numericSummaries.find((s) => s.column === 'Total Amount');

      expect(qty?.sum).toBe(6);
      expect(tax?.sum).toBe(63);
      expect(total?.sum).toBe(413);

      const status = overview.categoricalBreakdowns.find((c) => c.column === 'Payment Status');
      expect(status).toBeDefined();
      const paid = status!.rows.find((r) => r.key === 'Paid');
      const pending = status!.rows.find((r) => r.key === 'Pending');
      expect(paid?.count).toBe(2);
      expect(paid?.amount).toBe(354);
      expect(pending?.count).toBe(1);
      expect(pending?.amount).toBe(59);

      expect(overview.topRanking?.name).toBe('Sigma Ltd');
      expect(overview.topRanking?.amount).toBe(236);
      expect(overview.topRanking?.categoryColumn).toBe('Supplier');
    });
  });

  describe('formatSheetOverviewMarkdown', () => {
    it('emits structured markdown without internal vocabulary or mode pitch', () => {
      const md = formatSheetOverviewMarkdown(overviewFor(PURCHASE_REGISTER));

      expect(md).toMatch(/\*\*Purchase Register overview\*\*/);
      expect(md).toMatch(/Payment Status/);
      expect(md).toMatch(/Paid/);
      expect(md).toMatch(/Pending/);
      expect(md).toMatch(/Largest supplier/i);
      expect(md).toMatch(/Sigma Ltd/);
      expect(md).toMatch(/Want a totals row/);

      expect(md).not.toMatch(/Intent\s*:/i);
      expect(md).not.toMatch(/detectedType/i);
      expect(md).not.toMatch(/switch to Action/i);
      expect(md).not.toMatch(/\b[A-Z]:\s*(number|text|date)\b/);
    });

    it('includes data-quality notes in plain language for text dates and empty remarks', () => {
      const md = formatSheetOverviewMarkdown(overviewFor(PURCHASE_REGISTER));
      expect(md).toMatch(/Data quality notes/);
      expect(md).toMatch(/text dates|Remarks is entirely empty/i);
    });
  });

  describe('sanitizeAskAnswer', () => {
    it('strips Intent, detectedType, and mode-switch pitch', () => {
      const dirty =
        'Intent: EXPLAIN\ndetectedType: purchase_register\nHere is a summary.\nSwitch to Action mode to apply sorts.';
      const clean = sanitizeAskAnswer(dirty);
      expect(clean).not.toMatch(/Intent\s*:/i);
      expect(clean).not.toMatch(/detectedType/i);
      expect(clean).not.toMatch(/switch to Action/i);
      expect(clean).toMatch(/Here is a summary/);
    });
  });
});
