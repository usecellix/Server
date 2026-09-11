import {
  extractPortalStatementGstin,
  flagCrossGstinRows,
  normalizeGstin,
  proposeClientGstinFromBooks,
  validateGstinFormat,
  validatePortalGstinMatchesClient,
} from './gstin-validator';
import { NormalizedRowWithClientGstinColumn } from './types';

const VALID_A = '27ABCDE1234F1Z5';
const VALID_B = '29XYZAB5678C1Z9';

function row(
  partial: Partial<NormalizedRowWithClientGstinColumn>,
): NormalizedRowWithClientGstinColumn {
  return {
    rowIndex: 1,
    counterpartyGstin: '27AAAAA0000A1Z5',
    documentType: 'invoice',
    invoiceNumber: 'INV1',
    invoiceNumberRaw: 'INV1',
    invoiceDate: '2026-04-01',
    taxableValue: 1000,
    igst: 0,
    cgst: 90,
    sgst: 90,
    cess: 0,
    ...partial,
  };
}

describe('gstin-validator', () => {
  describe('normalizeGstin / validateGstinFormat', () => {
    it('normalizes whitespace and case', () => {
      expect(normalizeGstin(' 27abcde1234f1z5 ')).toBe(VALID_A);
    });

    it('accepts valid GSTIN', () => {
      const r = validateGstinFormat(VALID_A);
      expect(r.ok).toBe(true);
      expect(r.normalizedGstin).toBe(VALID_A);
    });

    it('rejects invalid format', () => {
      const r = validateGstinFormat('NOT-A-GSTIN');
      expect(r.ok).toBe(false);
      expect(r.blockingError).toMatch(/not a valid GSTIN/i);
    });

    it('rejects empty', () => {
      expect(validateGstinFormat('').ok).toBe(false);
    });
  });

  describe('validatePortalGstinMatchesClient', () => {
    it('passes when GSTINs match', () => {
      const r = validatePortalGstinMatchesClient(VALID_A, VALID_A.toLowerCase());
      expect(r.ok).toBe(true);
    });

    it('hard-stops on mismatch', () => {
      const r = validatePortalGstinMatchesClient(VALID_A, VALID_B);
      expect(r.ok).toBe(false);
      expect(r.blockingError).toMatch(/does not match/);
      expect(r.blockingError).toContain(VALID_A);
      expect(r.blockingError).toContain(VALID_B);
    });

    it('fails when portal GSTIN missing', () => {
      const r = validatePortalGstinMatchesClient(VALID_A, '');
      expect(r.ok).toBe(false);
    });
  });

  describe('flagCrossGstinRows', () => {
    it('splits clean vs cross-GSTIN rows from multi-branch sheet', () => {
      const rows = [
        row({ rowIndex: 1, clientSideGstin: VALID_A }),
        row({ rowIndex: 2, clientSideGstin: VALID_B }),
        row({ rowIndex: 3, clientSideGstin: VALID_A }),
        row({ rowIndex: 4, clientSideGstin: null }),
      ];
      const { clean, crossGstin } = flagCrossGstinRows(rows, VALID_A);
      expect(clean.map((r) => r.rowIndex)).toEqual([1, 3, 4]);
      expect(crossGstin.map((r) => r.rowIndex)).toEqual([2]);
    });
  });

  describe('extractPortalStatementGstin', () => {
    it('reads GSTIN from labeled header cell', () => {
      const grid = [
        ['GSTIN of the Taxpayer', VALID_A],
        ['GSTIN of supplier', 'Invoice number'],
        [VALID_B, 'INV-1'],
      ];
      expect(extractPortalStatementGstin(grid, 'purchase_vs_2b')).toBe(VALID_A);
    });

    it('falls back to first valid GSTIN in top rows', () => {
      const grid = [
        ['Report'],
        [VALID_B],
        ['GSTIN of supplier', 'Invoice'],
      ];
      expect(extractPortalStatementGstin(grid)).toBe(VALID_B);
    });
  });

  describe('proposeClientGstinFromBooks', () => {
    it('returns GSTIN when unambiguous', () => {
      expect(
        proposeClientGstinFromBooks([
          { clientSideGstin: VALID_A },
          { clientSideGstin: VALID_A },
        ]),
      ).toBe(VALID_A);
    });

    it('returns undefined when mixed', () => {
      expect(
        proposeClientGstinFromBooks([
          { clientSideGstin: VALID_A },
          { clientSideGstin: VALID_B },
        ]),
      ).toBeUndefined();
    });
  });
});
