import { FormattingChecker } from '../src/agents/checkers/formatting.checker';
import { WorkbookContext } from '../src/agents/types/agent.types';

describe('FormattingChecker', () => {
  const checker = new FormattingChecker();

  const context: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:C3',
        rowCount: 3,
        columnCount: 3,
        values: [
          ['Name', 'Date', 'Amount'],
          ['Row 1', '2024-01-01', 100],
          ['Row 2', '2024-02-01', 200],
        ],
        formulas: [['', '', ''], ['', '', ''], ['', '', '']],
        numberFormats: [
          ['General', 'd/m/yyyy', '₹#,##0.00'],
          ['General', 'd/m/yyyy', '₹#,##0.00'],
          ['General', 'd/m/yyyy', '₹#,##0.00'],
        ],
        structure: 'data_table',
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('flags explicit General format when adjacent row uses a custom format', () => {
    const result = checker.check(
      [
        {
          subtask: {
            id: 's1',
            description: 'Add row',
            targetSheet: 'Sheet1',
            dependsOn: [],
            estimatedActions: 1,
          },
          actions: [
            {
              type: 'ADD_ROW',
              data: ['Row 3', '2024-03-01', 300],
              format: { numberFormat: 'General' },
            },
          ],
        },
      ],
      context,
    );

    expect(result.passed).toBe(false);
    expect(result.issues.some((issue) => issue.severity === 'error')).toBe(true);
  });
});
