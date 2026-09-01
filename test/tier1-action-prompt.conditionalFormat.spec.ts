import { buildTier1UserMessage } from '../src/excel-ai/prompts/tier1-action-prompt';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #38 — Tier 1's user message must surface existing conditional-
 * format rules on the active sheet so the model can target one by id
 * (`existingRuleId`) instead of stacking a duplicate.
 */
describe('buildTier1UserMessage — existing conditional-format rules', () => {
  const baseContext: WorkbookContext = {
    activeSheetName: 'Purchase Register',
    sheets: [
      {
        name: 'Purchase Register',
        usedRange: 'A1:B3',
        rowCount: 3,
        columnCount: 2,
        values: [['Item', 'Total'], ['A', 1000], ['B', 2000]],
        formulas: [],
        numberFormats: [],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('lists existing rules on the active sheet for the CONDITIONAL_FORMAT hint', () => {
    const context: WorkbookContext = {
      ...baseContext,
      conditionalFormats: [
        {
          id: 'cf-1',
          sheetName: 'Purchase Register',
          range: 'B2:B3',
          ruleKind: 'cellValue',
          summary: 'greaterThan 1000',
        },
      ],
    };

    const result = buildTier1UserMessage('change the threshold to 1500', 'CONDITIONAL_FORMAT', context);
    expect(result).toContain(
      'Existing conditional-format rules on this sheet: [cf-1] B2:B3 (cellValue: greaterThan 1000)',
    );
  });

  it('says "(none)" when the sheet has no existing rules', () => {
    const context: WorkbookContext = { ...baseContext, conditionalFormats: [] };
    const result = buildTier1UserMessage('highlight expenses above 1000', 'CONDITIONAL_FORMAT', context);
    expect(result).toContain('Existing conditional-format rules on this sheet: (none)');
  });

  it('filters out rules belonging to a different sheet', () => {
    const context: WorkbookContext = {
      ...baseContext,
      conditionalFormats: [
        { id: 'cf-other', sheetName: 'Other Sheet', range: 'A1:A5', ruleKind: 'formula', summary: '=A1>0' },
      ],
    };
    const result = buildTier1UserMessage('highlight expenses above 1000', 'CONDITIONAL_FORMAT', context);
    expect(result).toContain('Existing conditional-format rules on this sheet: (none)');
    expect(result).not.toContain('cf-other');
  });

  it('omits the existing-rules line entirely for a non-CONDITIONAL_FORMAT hint', () => {
    const context: WorkbookContext = {
      ...baseContext,
      conditionalFormats: [
        { id: 'cf-1', sheetName: 'Purchase Register', range: 'B2:B3', ruleKind: 'cellValue', summary: 'x' },
      ],
    };
    const result = buildTier1UserMessage('sort by total', 'SORT_OR_FILTER', context);
    expect(result).not.toContain('Existing conditional-format rules');
  });

  it('handles an undefined conditionalFormats field (older/minimal contexts) as "(none)"', () => {
    const result = buildTier1UserMessage('highlight expenses above 1000', 'CONDITIONAL_FORMAT', baseContext);
    expect(result).toContain('Existing conditional-format rules on this sheet: (none)');
  });
});
