import {
  annotateExplicitOverwriteConfirmation,
  hasExplicitOverwriteConfirmation,
  isRefinementOfOwnLastEdit,
  rangesOverlap,
  resolveActionWriteTarget,
} from '../src/excel-ai/utils/overwrite-confirmation.util';
import {
  collectRecentTurnActionRecords,
  extractTurnActionRecords,
} from '../src/excel-ai/utils/turn-action-history.util';

describe('Spec 21 overwrite confirmation', () => {
  describe('hasExplicitOverwriteConfirmation', () => {
    it('matches change/update/replace/correct phrasing', () => {
      expect(hasExplicitOverwriteConfirmation('change to paid invoices')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('update the remarks to Paid')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('replace with Y')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('correct the values to X')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('modify remarks as Pending')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('fix those cells to blank')).toBe(true);
      expect(hasExplicitOverwriteConfirmation('overwrite the column with Paid')).toBe(true);
    });

    it('does not match plain add/set phrasing', () => {
      expect(hasExplicitOverwriteConfirmation('add remarks to paid invoices')).toBe(false);
      expect(hasExplicitOverwriteConfirmation('set remarks for paid rows')).toBe(false);
      expect(hasExplicitOverwriteConfirmation('add a column called Net of Tax')).toBe(false);
      expect(hasExplicitOverwriteConfirmation('write Paid into Remarks')).toBe(false);
    });
  });

  describe('rangesOverlap + isRefinementOfOwnLastEdit', () => {
    it('detects overlap when follow-up range is a subset of prior write', () => {
      expect(
        rangesOverlap('Purchase Register!L2:L51', 'Purchase Register!L2:L10'),
      ).toBe(true);
      expect(
        rangesOverlap('Purchase Register!A1:L51', 'Purchase Register!L2:L51'),
      ).toBe(true);
      expect(rangesOverlap('Sheet1!A1:A10', 'Sheet1!B1:B10')).toBe(false);
      expect(rangesOverlap('Sheet1!A1:A10', 'Other!A1:A10')).toBe(false);
    });

    it('treats prior-turn overlap as refinement even without restating the range', () => {
      const priorColumnScoped = [
        {
          actionType: 'SET_MATCHING_ROWS',
          sheetName: 'Purchase Register',
          affectedRange: 'Purchase Register!A1:L51',
          targetColumn: 'Remarks',
          turnIndex: 0,
        },
      ];

      expect(
        isRefinementOfOwnLastEdit('Purchase Register!L2:L51', priorColumnScoped, {
          sheetName: 'Purchase Register',
          targetColumn: 'Remarks',
        }),
      ).toBe(true);

      // Same named column without restating an exact A1 range.
      expect(
        isRefinementOfOwnLastEdit('Purchase Register!A1:L51', priorColumnScoped, {
          sheetName: 'Purchase Register',
          targetColumn: 'Remarks',
        }),
      ).toBe(true);

      // Different column must not ride on the wide table range.
      expect(
        isRefinementOfOwnLastEdit('Purchase Register!K2', priorColumnScoped, {
          sheetName: 'Purchase Register',
          targetColumn: 'Payment Status',
        }),
      ).toBe(false);

      // Precise prior cell/column range (no targetColumn) still overlaps geometrically.
      const priorPrecise = [
        {
          actionType: 'SET_CELL',
          sheetName: 'Purchase Register',
          affectedRange: 'Purchase Register!L2:L51',
          turnIndex: 0,
        },
      ];
      expect(isRefinementOfOwnLastEdit('Purchase Register!L2', priorPrecise)).toBe(true);
      expect(isRefinementOfOwnLastEdit('Purchase Register!K2', priorPrecise)).toBe(false);
    });
  });

  describe('annotateExplicitOverwriteConfirmation', () => {
    it('Turn 1 → Turn 2 repro: change-to language confirms overwrite', () => {
      const actions = [
        {
          type: 'SET_MATCHING_ROWS' as const,
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          hasHeaders: true,
          filter: { column: 'Payment Status', operator: 'equals' as const, value: 'Paid' },
          targetColumn: 'Remarks',
          value: 'paid invoices',
        },
      ];

      const annotated = annotateExplicitOverwriteConfirmation(
        actions,
        'change to paid invoices',
        [],
      );
      expect(annotated[0]?.explicitOverwriteConfirmed).toBe(true);
    });

    it('confirms via prior-turn overlap without change language', () => {
      const prior = extractTurnActionRecords([
        {
          type: 'SET_MATCHING_ROWS',
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          hasHeaders: true,
          filter: { column: 'Payment Status', operator: 'equals', value: 'Paid' },
          targetColumn: 'Remarks',
          value: 'Paid',
        },
      ]);
      expect(prior[0]?.affectedRange).toContain('A1:L51');

      const followUp = [
        {
          type: 'SET_MATCHING_ROWS' as const,
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          hasHeaders: true,
          filter: { column: 'Payment Status', operator: 'equals' as const, value: 'Paid' },
          targetColumn: 'Remarks',
          value: 'paid invoices',
        },
      ];

      // No change/update verb — refinement of own last edit still confirms.
      const annotated = annotateExplicitOverwriteConfirmation(
        followUp,
        'put paid invoices in remarks for those rows',
        prior,
      );
      expect(annotated[0]?.explicitOverwriteConfirmed).toBe(true);
    });

    it('does not confirm Net of Tax Bug 6 (no prior context, no change language)', () => {
      const actions = [
        {
          type: 'SET_FORMULA' as const,
          sheetName: 'Purchase Register',
          row: 1,
          col: 10,
          formula: '=J2-I2',
        },
      ];

      const annotated = annotateExplicitOverwriteConfirmation(
        actions,
        'Add a column called Net of Tax that subtracts Tax Amount from Total Amount',
        [],
      );
      expect(annotated[0]?.explicitOverwriteConfirmed).toBeFalsy();
    });

    it('does not confirm plain add remarks without prior turn overlap', () => {
      const actions = [
        {
          type: 'SET_MATCHING_ROWS' as const,
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          targetColumn: 'Remarks',
          value: 'Paid',
        },
      ];
      const annotated = annotateExplicitOverwriteConfirmation(
        actions,
        'add remarks to paid invoices',
        [],
      );
      expect(annotated[0]?.explicitOverwriteConfirmed).toBeFalsy();
    });
  });

  describe('turn history write extraction', () => {
    it('records affectedRange for SET_MATCHING_ROWS and collects from conversation', () => {
      const records = extractTurnActionRecords([
        {
          type: 'SET_MATCHING_ROWS',
          sheetName: 'Purchase Register',
          range: 'A1:L51',
          targetColumn: 'Remarks',
          value: 'Paid',
        },
      ]);
      expect(records).toEqual([
        expect.objectContaining({
          actionType: 'SET_MATCHING_ROWS',
          sheetName: 'Purchase Register',
          affectedRange: 'Purchase Register!A1:L51',
          targetColumn: 'Remarks',
        }),
      ]);

      const collected = collectRecentTurnActionRecords([
        {
          role: 'assistant',
          metadata: { turnActionRecords: records },
        },
      ]);
      expect(collected).toHaveLength(1);
      expect(collected[0].affectedRange).toBe('Purchase Register!A1:L51');
    });

    it('resolveActionWriteTarget maps SET_FORMULA row/col to K2', () => {
      expect(
        resolveActionWriteTarget({
          type: 'SET_FORMULA',
          sheetName: 'Purchase Register',
          row: 1,
          col: 10,
        }),
      ).toEqual({
        sheetName: 'Purchase Register',
        affectedRange: 'Purchase Register!K2',
      });
    });
  });
});
