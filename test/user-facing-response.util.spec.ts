import { CellChange } from '../src/audit/types/change-set.types';
import { SheetAction } from '../src/excel-ai/types/sheet-actions.types';
import {
  buildInternalDetails,
  buildUserFacingSummary,
  describeRangeCompactly,
  INTERNAL_COPY_MARKERS,
  sanitizeAnswerForHeadline,
  tierProcessingLabel,
} from '../src/excel-ai/utils/user-facing-response.util';

function cell(sheet: string, address: string): CellChange {
  return {
    cell: address,
    sheet,
    before: null,
    after: 1,
    isHardcoded: true,
  };
}

describe('user-facing-response.util', () => {
  describe('describeRangeCompactly', () => {
    it('collapses a rectangular block to a compact A1 range', () => {
      const changes = [
        cell('Purchase Register', 'A53'),
        cell('Purchase Register', 'B53'),
        cell('Purchase Register', 'C53'),
        cell('Purchase Register', 'A54'),
        cell('Purchase Register', 'B54'),
        cell('Purchase Register', 'C54'),
        cell('Purchase Register', 'A55'),
        cell('Purchase Register', 'B55'),
        cell('Purchase Register', 'C55'),
      ];
      expect(describeRangeCompactly(changes)).toBe('Purchase Register!A53:C55');
    });
  });

  describe('buildUserFacingSummary — 9-cell / 1-action regression', () => {
    it('derives supporting detail from cell diffs, not action count', () => {
      const actions: SheetAction[] = [
        {
          type: 'AGGREGATE_TABLE',
          sheetName: 'Purchase Register',
          groupByColumn: 'Payment Status',
        },
      ];
      const changes = [
        cell('Purchase Register', 'A53'),
        cell('Purchase Register', 'B53'),
        cell('Purchase Register', 'C53'),
        cell('Purchase Register', 'A54'),
        cell('Purchase Register', 'B54'),
        cell('Purchase Register', 'C54'),
        cell('Purchase Register', 'A55'),
        cell('Purchase Register', 'B55'),
        cell('Purchase Register', 'C55'),
      ];

      const summary = buildUserFacingSummary({
        answer: 'Added a Payment Status summary.',
        actions,
        changes,
        activeSheetName: 'Purchase Register',
      });

      expect(summary.contextLine).toBe('Working with: Purchase Register');
      expect(summary.headline).toBe('Added a Payment Status summary.');
      expect(summary.supportingDetail).toBe('9 cells, Purchase Register!A53:C55');
      expect(summary.bullets).toBeUndefined();
      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
    });
  });

  describe('assumption / typo fixture', () => {
    it('states the assumption in the visible headline', () => {
      const summary = buildUserFacingSummary({
        answer: "I'll sort with Paid first, then Pending.",
        actions: [{ type: 'SORT_RANGE', sheetName: 'Purchase Register', key: 10, ascending: true }],
        assumption:
          "I noticed 'paid should be first then paid' looks like it may have a typo — I've sorted with Paid first, then Pending",
        activeSheetName: 'Purchase Register',
      });

      expect(summary.headline.toLowerCase()).toContain('typo');
      expect(summary.headline.toLowerCase()).toContain('paid first');
      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
    });
  });

  describe('internal copy never in headline', () => {
    it('strips tier jargon from answer and falls back to action description', () => {
      expect(
        sanitizeAnswerForHeadline(
          'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
        ),
      ).toBe('');

      const summary = buildUserFacingSummary({
        answer: 'Tier 1 single-action (CONDITIONAL_FORMAT) — one LLM call, no verification.',
        actions: [
          {
            type: 'FORMAT_MATCHING_ROWS',
            sheetName: 'Purchase Register',
            format: { fillColor: '#FFEB9C' },
          },
        ],
      });

      expect(INTERNAL_COPY_MARKERS.test(summary.headline)).toBe(false);
      expect(summary.headline.toLowerCase()).toContain('highlight');
    });
  });

  describe('buildInternalDetails', () => {
    it('preserves legacy tier explanation and raw actions', () => {
      const details = buildInternalDetails({
        tier: 1,
        model: 'openai/gpt-5-mini',
        processingLabel: tierProcessingLabel(1, 'CONDITIONAL_FORMAT'),
        actions: [{ type: 'FORMAT_MATCHING_ROWS', range: 'A2:L51', sheetName: 'Purchase Register' }],
        legacyExplanation: tierProcessingLabel(1, 'CONDITIONAL_FORMAT'),
        assumption: 'Paid first, then Pending',
      });

      expect(details.processingLabel).toContain('Tier 1');
      expect(details.model).toBe('openai/gpt-5-mini');
      expect(details.rawActionSummary).toContain('FORMAT_MATCHING_ROWS');
      expect(details.assumption).toContain('Paid first');
    });
  });
});
