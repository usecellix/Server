import * as fs from 'fs';
import * as path from 'path';
import { GstReconService } from './gst-recon.service';

interface Fixture {
  booksSheet: string;
  portalSheet: string;
  booksGrid: unknown[][];
  portalGrid: unknown[][];
  portalHeadersRow: number;
  expected: {
    totalBooksRows: number;
    matched: number;
    blank_counterparty_gstin: number;
    blank_taxable_value: number;
    amount_mismatch: number;
    date_mismatch: number;
    genuinely_missing: number;
    gstin_not_in_portal: number;
  };
}

function loadFixture(name: string): Fixture {
  const p = path.join(__dirname, '__fixtures__', name);
  return JSON.parse(fs.readFileSync(p, 'utf8')) as Fixture;
}

/**
 * Permanent regression fixture for the production bug this whole patch was written for:
 * a books export with no Invoice Number column at all, reported as 0 matched / everything
 * dumped into a flat unmatched pile. This exercises the full real path — parsing, Pass 1b,
 * and per-reason diagnosis — together, not in isolation.
 */
describe('no-invoice-number-column fixture — full missed_books_only path', () => {
  const service = new GstReconService();
  const fx = loadFixture('no-invoice-number-column.json');

  it('matches a non-zero count via Pass 1b and accounts for every row across matched + reasons', async () => {
    const result = await service.reconcile({
      reconciliation_type: 'PR_VS_GSTR2B',
      missed_books_only: true,
      purchase_register: {
        sheet_name: fx.booksSheet,
        data: fx.booksGrid,
        headers_row: 1,
      },
      portal_file: {
        sheet_name: fx.portalSheet,
        data: fx.portalGrid,
        headers_row: fx.portalHeadersRow,
        file_type: 'GSTR2B',
      },
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);

    const s = result.summary;

    // Non-zero matched — the exact regression this patch fixes (was 0 in production).
    expect(s.exact_matched).toBeGreaterThan(0);
    expect(s.exact_matched).toBe(fx.expected.matched);

    // Every row accounted for: matched + diagnosed reasons === total books rows, none dropped.
    const totalAccounted =
      s.exact_matched +
      s.mismatch_blank_gstin +
      s.mismatch_blank_taxable_value +
      s.mismatch_amount +
      s.mismatch_date +
      s.mismatch_genuinely_missing;
    expect(totalAccounted).toBe(fx.expected.totalBooksRows);
    expect(s.pr_only).toBe(fx.expected.totalBooksRows - fx.expected.matched);

    // Specific reason breakdown, not just totals.
    expect(s.mismatch_blank_gstin).toBe(fx.expected.blank_counterparty_gstin);
    expect(s.mismatch_blank_taxable_value).toBe(fx.expected.blank_taxable_value);
    expect(s.mismatch_amount).toBe(fx.expected.amount_mismatch);
    expect(s.mismatch_date).toBe(fx.expected.date_mismatch);
    // gstin_not_in_portal and genuinely_missing are combined in the reporting bucket.
    expect(s.mismatch_genuinely_missing).toBe(
      fx.expected.genuinely_missing + fx.expected.gstin_not_in_portal,
    );

    // The grouped sheet was actually produced with a section per reason.
    const writeTable = (
      result.actions as Array<{ type: string; headers?: string[]; rows?: unknown[][] }>
    ).find((a) => a.type === 'WRITE_TABLE');
    expect(writeTable).toBeDefined();
    const sectionTitles = writeTable!.rows!.map((r) => r[0]).filter((c) => typeof c === 'string');
    expect(sectionTitles).toEqual(
      expect.arrayContaining([
        'Blank GSTIN rows',
        'Blank taxable value rows',
        'Amount mismatch rows',
        'Date mismatch rows',
        'Genuinely missing rows',
      ]),
    );
    // Nothing lands in the defensive "unclassified" catch-all for this fixture.
    expect(sectionTitles.some((t) => String(t).includes('Unclassified'))).toBe(false);
  });
});
