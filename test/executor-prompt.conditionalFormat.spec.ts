import { buildExecutorUserMessage } from '../src/agents/prompts/executor.prompt';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #38 — the Executor's user message must surface existing
 * conditional-format rules on the target sheet so it can target one by id
 * (`existingRuleId`) instead of stacking a duplicate.
 */
describe('buildExecutorUserMessage — existing conditional-format rules', () => {
  const subtask: SubTask = {
    id: 's1',
    description: 'Change the highlight threshold to 1500',
    targetSheet: 'Purchase Register',
    dependsOn: [],
    estimatedActions: 1,
  };

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

  it('lists existing rules scoped to the subtask target sheet', () => {
    const context: WorkbookContext = {
      ...baseContext,
      conditionalFormats: [
        { id: 'cf-1', sheetName: 'Purchase Register', range: 'B2:B3', ruleKind: 'cellValue', summary: 'greaterThan 1000' },
      ],
    };

    const result = buildExecutorUserMessage(subtask, context, []);
    expect(result).toContain('Existing conditional-format rules on this sheet:');
    expect(result).toContain('- [cf-1] B2:B3 (cellValue: greaterThan 1000)');
  });

  it('filters out rules belonging to a different sheet', () => {
    const context: WorkbookContext = {
      ...baseContext,
      conditionalFormats: [
        { id: 'cf-other', sheetName: 'Other Sheet', range: 'A1:A5', ruleKind: 'formula', summary: '=A1>0' },
      ],
    };
    const result = buildExecutorUserMessage(subtask, context, []);
    expect(result).not.toContain('Existing conditional-format rules');
    expect(result).not.toContain('cf-other');
  });

  it('renders nothing when there are no conditional-format rules at all', () => {
    const result = buildExecutorUserMessage(subtask, baseContext, []);
    expect(result).not.toContain('Existing conditional-format rules');
  });
});
