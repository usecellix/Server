import {
  ensureNumberFormatPlanSafety,
  rebindFormatRangeNumberFormats,
} from '../src/agents/utils/preserve-number-format.util';
import type { Action, PlannerOutput, WorkbookContext } from '../src/agents/types/agent.types';

const sheetWithDateFormat = (format: string): WorkbookContext => ({
  activeSheetName: 'Purchase Register',
  sheets: [
    {
      name: 'Purchase Register',
      usedRange: 'A1:B3',
      rowCount: 3,
      columnCount: 2,
      values: [
        ['Date', 'Amount'],
        ['01/15/2024', 100],
        ['02/20/2024', 200],
      ],
      formulas: [
        ['', ''],
        ['', ''],
        ['', ''],
      ],
      numberFormats: [
        ['General', 'General'],
        [format, '₹#,##0.00'],
        [format, '₹#,##0.00'],
      ],
      structure: 'data_table',
      headerRowIndex: 0,
    },
  ],
  namedRanges: [],
  tables: [],
});

describe('preserve-number-format.util', () => {
  describe('rebindFormatRangeNumberFormats', () => {
    it('replaces invented dd-mm-yyyy with dominant column format', () => {
      const actions: Action[] = [
        {
          type: 'FORMAT_RANGE',
          sheetName: 'Purchase Register',
          row: 1,
          col: 0,
          rowCount: 2,
          colCount: 1,
          format: { numberFormat: 'dd-mm-yyyy' },
        },
      ];

      const { actions: out, droppedInvented } = rebindFormatRangeNumberFormats(actions, {
        userPrompt: 'change the date back to the original format',
        context: sheetWithDateFormat('m/d/yyyy'),
      });

      expect(droppedInvented).toBe(false);
      expect(out[0]?.format?.numberFormat).toBe('m/d/yyyy');
    });

    it('drops invented date format when column has no format to re-apply', () => {
      const actions: Action[] = [
        {
          type: 'FORMAT_RANGE',
          sheetName: 'Purchase Register',
          row: 1,
          col: 0,
          rowCount: 2,
          colCount: 1,
          format: { numberFormat: 'dd-mm-yyyy' },
        },
      ];

      const { actions: out, droppedInvented } = rebindFormatRangeNumberFormats(actions, {
        userPrompt: 'change the date back to the original format',
        context: sheetWithDateFormat('General'),
      });

      expect(droppedInvented).toBe(true);
      expect(out).toHaveLength(0);
    });

    it('keeps format the user explicitly named', () => {
      const actions: Action[] = [
        {
          type: 'FORMAT_RANGE',
          sheetName: 'Purchase Register',
          row: 1,
          col: 0,
          rowCount: 2,
          colCount: 1,
          format: { numberFormat: 'dd-mm-yyyy' },
        },
      ];

      const { actions: out } = rebindFormatRangeNumberFormats(actions, {
        userPrompt: 'format dates as dd-mm-yyyy',
        context: sheetWithDateFormat('m/d/yyyy'),
      });

      expect(out[0]?.format?.numberFormat).toBe('dd-mm-yyyy');
    });
  });

  describe('ensureNumberFormatPlanSafety', () => {
    it('blocks plans that assume dd-mm-yyyy without user naming it', () => {
      const plan: PlannerOutput = {
        subtasks: [
          {
            id: 's1',
            description:
              "Reapply the sheet's original date format (assumed dd-mm-yyyy) to Date column",
            targetSheet: 'Purchase Register',
            dependsOn: [],
            estimatedActions: 1,
          },
        ],
        clarificationsNeeded: [],
        confidence: 'medium',
        reasoning: 'Assumed Indian date',
      };

      const safe = ensureNumberFormatPlanSafety(
        'change the date back to the original format',
        plan,
      );
      expect(safe.subtasks).toHaveLength(0);
      expect(safe.clarificationsNeeded.length).toBeGreaterThan(0);
      expect(safe.clarificationsNeeded.join(' ')).not.toMatch(/switch to Action/i);
    });
  });
});
