import { classifyComplexity } from '../src/excel-ai/utils/complexity-classifier.util';
import { buildRefinementContext } from '../src/excel-ai/utils/refinement-context.util';
import { ChangeSetRecord } from '../src/audit/types/change-set.types';

describe('write route tier classification', () => {
  it('classifies a multi-step parent request as tier 3 compound', () => {
    const parent = classifyComplexity('sort by column B and then create a chart');
    expect(parent.match?.tier).toBe(3);
  });

  it('classifies a quick-edit follow-up independently as tier 1', () => {
    const quickEdit = classifyComplexity('sort column B descending by value');
    expect(quickEdit.match?.tier).toBe(1);
    expect(quickEdit.match?.actionHint).toBe('SORT_OR_FILTER');
  });

  it('does not attach tier metadata to refinement context payloads', () => {
    const changeSet = {
      changeSetId: 'cs-1',
      conversationId: 'conv-1',
      traceId: 'trace-1',
      timestamp: new Date(),
      prompt: 'reconcile bank statement across all sheets',
      beforeState: {},
      changes: [
        {
          cell: 'B2',
          sheet: 'Sheet1',
          before: 100,
          after: 200,
          isHardcoded: false,
        },
      ],
      actions: [],
      status: 'applied',
    } as ChangeSetRecord;

    const refinement = buildRefinementContext(changeSet);
    expect(refinement).not.toHaveProperty('tier');
    expect(refinement).not.toHaveProperty('complexity');
    expect(refinement.promptContext.length).toBeGreaterThan(0);
  });
});
