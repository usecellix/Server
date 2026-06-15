import { buildRefinementContext } from '../src/excel-ai/utils/refinement-context.util';
import { ChangeSetRecord } from '../src/audit/types/change-set.types';

describe('buildRefinementContext', () => {
  it('builds sparse sheet data and prompt context from a change set', () => {
    const changeSet: ChangeSetRecord = {
      changeSetId: 'cs_1',
      conversationId: 'conv_1',
      traceId: 'trace_1',
      timestamp: new Date('2024-01-01T00:00:00Z'),
      prompt: 'Add GST row',
      beforeState: {
        'Sheet1!A1': { value: 'Name', formula: '', format: 'General' },
        'Sheet1!B1': { value: 'Amount', formula: '', format: 'General' },
        'Sheet1!A2': { value: 'Row 1', formula: '', format: 'General' },
        'Sheet1!B2': { value: 100, formula: '', format: 'General' },
      },
      changes: [
        {
          cell: 'A3',
          sheet: 'Sheet1',
          before: '',
          after: 'GST',
          isHardcoded: true,
        },
        {
          cell: 'B3',
          sheet: 'Sheet1',
          before: '',
          after: '=B2*0.18',
          formula: '=B2*0.18',
          isHardcoded: false,
        },
      ],
      actions: [],
      status: 'applied',
    };

    const result = buildRefinementContext(changeSet);

    expect(result.sheetData.length).toBeGreaterThan(0);
    expect(result.richWorkbookContext.activeSheet).toBe('Sheet1');
    expect(result.promptContext).toContain('QUICK EDIT MODE');
    expect(result.promptContext).toContain('Add GST row');
  });
});
