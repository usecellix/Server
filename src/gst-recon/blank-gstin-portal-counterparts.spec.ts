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
    originalBlankGstinCount: number;
    blank_gstin_likely_matched: number;
    blank_counterparty_gstin_remaining: number;
    portal_only_after_fix: number;
  };
}

function loadFixture(name: string): Fixture {
  const p = path.join(__dirname, '__fixtures__', name);
  return JSON.parse(fs.readFileSync(p, 'utf8')) as Fixture;
}

/**
 * Permanent regression fixture for the blank-GSTIN double-counting bug: 9 blank-GSTIN
 * books rows whose real counterpart invoice exists in the portal file under a distinct
 * GSTIN, plus 1 genuinely unresolvable blank-GSTIN row and 1 genuinely unrelated
 * portal-only row (to prove the fix doesn't over-match). Exercises the full real path —
 * parsing, diagnosis, and the blank-GSTIN resolution pass together.
 */
describe('blank-gstin-portal-counterparts fixture — full missed_books_only path', () => {
  const service = new GstReconService();
  const fx = loadFixture('blank-gstin-portal-counterparts.json');

  it('reclassifies the 9 resolvable blank-GSTIN rows and removes their portal_only counterparts, without double-counting', async () => {
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

    expect(s.total_pr_rows).toBe(fx.expected.totalBooksRows);
    expect(s.mismatch_blank_gstin_likely_matched).toBe(fx.expected.blank_gstin_likely_matched);
    expect(s.mismatch_blank_gstin).toBe(fx.expected.blank_counterparty_gstin_remaining);
    expect(s.portal_only).toBe(fx.expected.portal_only_after_fix);

    // Invariant: likely-matched + remaining blank-GSTIN === original blank-GSTIN count.
    expect(s.mismatch_blank_gstin_likely_matched + s.mismatch_blank_gstin).toBe(
      fx.expected.originalBlankGstinCount,
    );

    // Every books row is accounted for: matched + all mismatch reasons === total books rows.
    const totalAccounted =
      s.exact_matched +
      s.mismatch_blank_gstin +
      s.mismatch_blank_gstin_likely_matched +
      s.mismatch_blank_taxable_value +
      s.mismatch_amount +
      s.mismatch_date +
      s.mismatch_genuinely_missing;
    expect(totalAccounted).toBe(fx.expected.totalBooksRows);

    // The grouped sheet has its own section for the likely-matched rows, with suggested GSTINs shown.
    const writeTable = (
      result.actions as Array<{ type: string; headers?: string[]; rows?: unknown[][] }>
    ).find((a) => a.type === 'WRITE_TABLE');
    expect(writeTable).toBeDefined();
    const rows = writeTable!.rows!;
    const sectionIdx = rows.findIndex(
      (r, i) =>
        r[0] === 'Blank GSTIN — likely matched (confirm & fill in GSTIN)' &&
        rows[i + 1]?.[0] === 'Vendor Name (Books)',
    );
    expect(sectionIdx).toBeGreaterThan(-1);
    const dataRows = rows
      .slice(sectionIdx + 2)
      .filter((r) => typeof r[0] === 'string' && (r[0] as string).startsWith('Vendor '));
    expect(dataRows).toHaveLength(9);
    // Every likely-matched row shows a suggested GSTIN (column index 3).
    expect(dataRows.every((r) => typeof r[3] === 'string' && (r[3] as string).startsWith('32AAAAA'))).toBe(
      true,
    );
  });
});
