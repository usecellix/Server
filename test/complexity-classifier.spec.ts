import * as fs from 'fs';
import * as path from 'path';
import { classifyComplexity } from '../src/excel-ai/utils/complexity-classifier.util';

const catalogFixtures = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'fixtures', 'catalog-classification.json'), 'utf8'),
) as Array<{
  message: string;
  expectedTier: number | null;
  expectedActionHint: string | null;
}>;

describe('classifyComplexity', () => {
  describe('compound signal priority', () => {
    it('escalates to tier 3 when compound signal matches alongside a single-action pattern', () => {
      const result = classifyComplexity('bold A1:C1 and then sort column B');
      expect(result.match).toEqual({
        tier: 3,
        actionHint: 'CELL_FORMAT',
        matchedBy: 'regex',
      });
    });

    it('escalates sort+chart compound to tier 3', () => {
      const result = classifyComplexity('sort by column B and then create a chart');
      expect(result.match?.tier).toBe(3);
      expect(result.match?.actionHint).toBe('SORT_OR_FILTER');
    });

    it('returns null when compound signal has no recognizable single action', () => {
      const result = classifyComplexity('do this and then do that other thing');
      expect(result.match).toBeNull();
    });

    it('does not treat simple comma lists as compound unless signal matches', () => {
      const result = classifyComplexity('bold A1, italic B1');
      expect(result.match?.tier).toBe(0);
      expect(result.match?.actionHint).toBe('CELL_FORMAT');
    });
  });

  describe('numeric/financial find-replace escalation', () => {
    it('escalates tier 1 FIND_REPLACE to tier 2 for GST targets', () => {
      const result = classifyComplexity('find and replace GST in column D');
      expect(result.match).toEqual({
        tier: 2,
        actionHint: 'FIND_REPLACE',
        matchedBy: 'regex',
      });
    });

    it('escalates for amount/total/invoice/tax hints', () => {
      expect(classifyComplexity('find and replace amount values in column E').match?.tier).toBe(2);
      expect(classifyComplexity('find and replace invoice numbers').match?.tier).toBe(2);
      expect(classifyComplexity('find and replace tax codes').match?.tier).toBe(2);
    });

    it('keeps non-financial find-replace at tier 1', () => {
      const result = classifyComplexity('find and replace ABC with XYZ');
      expect(result.match).toEqual({
        tier: 1,
        actionHint: 'FIND_REPLACE',
        matchedBy: 'regex',
      });
    });
  });

  describe('tier 0 patterns', () => {
    it('matches CELL_FORMAT with explicit cell reference', () => {
      expect(classifyComplexity('bold cells A1 to C1').match).toEqual({
        tier: 0,
        actionHint: 'CELL_FORMAT',
        matchedBy: 'regex',
      });
    });

    it('matches FREEZE_PANES', () => {
      expect(classifyComplexity('freeze top row').match?.actionHint).toBe('FREEZE_PANES');
    });

    it('matches VISIBILITY_TOGGLE', () => {
      expect(classifyComplexity('hide column F').match?.actionHint).toBe('VISIBILITY_TOGGLE');
    });

    it('matches ROW_COL_STRUCTURE', () => {
      expect(classifyComplexity('insert a row above row 10').match?.actionHint).toBe('ROW_COL_STRUCTURE');
    });
  });

  describe('tier 1 patterns', () => {
    it('matches SORT_OR_FILTER', () => {
      expect(classifyComplexity('sort column B descending by value').match).toEqual({
        tier: 1,
        actionHint: 'SORT_OR_FILTER',
        matchedBy: 'regex',
      });
    });

    it('matches CONDITIONAL_FORMAT', () => {
      expect(classifyComplexity('highlight cells greater than 100').match?.actionHint).toBe(
        'CONDITIONAL_FORMAT',
      );
      expect(classifyComplexity('remvoe the highlights red').match).toEqual({
        tier: 1,
        actionHint: 'CONDITIONAL_FORMAT',
        matchedBy: 'regex',
      });
    });

    it('matches COPY_FILL', () => {
      expect(classifyComplexity('fill down the formula in column D').match?.actionHint).toBe('COPY_FILL');
    });
  });

  describe('tier 2 patterns', () => {
    it('matches FORMULA_GEN', () => {
      expect(classifyComplexity('calculate GST at 18% for column D').match?.actionHint).toBe('FORMULA_GEN');
    });

    it('matches PIVOT_TABLE', () => {
      expect(classifyComplexity('create a pivot table from this data').match?.actionHint).toBe('PIVOT_TABLE');
    });

    it('matches CHART', () => {
      expect(classifyComplexity('create a chart from sales data').match?.actionHint).toBe('CHART');
      expect(classifyComplexity('create charts from sales data').match?.actionHint).toBe('CHART');
    });

    it('matches DUPLICATE_CHECK', () => {
      expect(classifyComplexity('find duplicate invoice numbers').match?.actionHint).toBe('DUPLICATE_CHECK');
    });

    it('matches DATA_VALIDATION', () => {
      expect(classifyComplexity('add data validation dropdown in column B').match?.actionHint).toBe(
        'DATA_VALIDATION',
      );
    });

    it('matches ERROR_FIX', () => {
      expect(classifyComplexity('fix #REF! error in cell G5').match?.actionHint).toBe('ERROR_FIX');
    });
  });

  describe('tier 3 patterns', () => {
    it('routes dashboards to the multi-step planner', () => {
      const match = classifyComplexity('Create a dashboard with charts and KPIs.').match;
      expect(match?.tier).toBe(3);
      expect(match?.actionHint).toBe('DASHBOARD');
    });
  });

  describe('no match', () => {
    it('returns null for ambiguous multi-step requests without compound regex signal', () => {
      expect(
        classifyComplexity('reconcile the bank statement against the ledger and flag mismatches').match,
      ).toBeNull();
    });

    it('returns null for generic help requests', () => {
      expect(classifyComplexity('help me understand this workbook').match).toBeNull();
    });
  });

  describe('catalog classification fixtures', () => {
    it.each(catalogFixtures)('classifies "$message"', ({ message, expectedTier, expectedActionHint }) => {
      const result = classifyComplexity(message);

      if (expectedTier === null) {
        expect(result.match).toBeNull();
        return;
      }

      expect(result.match?.tier).toBe(expectedTier);
      if (expectedActionHint === null) {
        expect(result.match?.actionHint).toBeDefined();
      } else {
        expect(result.match?.actionHint).toBe(expectedActionHint);
      }
    });
  });
});
