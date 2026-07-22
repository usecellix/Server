import { Tier1SingleActionService } from '../src/excel-ai/services/tier1-single-action.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { classifyComplexity } from '../src/excel-ai/utils/complexity-classifier.util';
import { WorkbookContext } from '../src/agents/types/agent.types';

const workbookContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:C20',
      rowCount: 20,
      columnCount: 3,
      values: [
        ['Name', 'Status', 'Notes'],
        ['Alpha', 'Open', ''],
        ['Beta', 'Closed', ''],
      ],
      formulas: [],
      numberFormats: [],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('Tier1SingleActionService', () => {
  let openRouter: jest.Mocked<Pick<OpenRouterService, 'complete'>>;
  let service: Tier1SingleActionService;

  beforeEach(() => {
    openRouter = { complete: jest.fn() };
    const config = {
      openRouterModelLow: 'openai/gpt-5-mini',
    } as unknown as AppConfigService;
    service = new Tier1SingleActionService(
      openRouter as unknown as OpenRouterService,
      config,
    );
  });

  it('makes exactly one LLM call per request', async () => {
    openRouter.complete.mockResolvedValue(
      JSON.stringify({
        answer: 'Sorted column B descending.',
        actions: [
          {
            type: 'SORT_RANGE',
            sheetName: 'Sheet1',
            range: 'A1:C20',
            key: 1,
            ascending: false,
            hasHeaders: true,
            columnName: 'Status',
          },
        ],
      }),
    );

    const result = await service.execute(
      'sort column B descending by value',
      'SORT_OR_FILTER',
      workbookContext,
    );

    expect(openRouter.complete).toHaveBeenCalledTimes(1);
    expect(openRouter.complete).toHaveBeenCalledWith(
      expect.objectContaining({ tier: 'low' }),
    );
    expect(result.actions).toHaveLength(1);
    expect(result.actions[0].type).toBe('SORT_RANGE');
  });

  it('rejects numeric/financial find-replace defensively', async () => {
    const classified = classifyComplexity('find and replace GST in column D');
    expect(classified.match?.tier).toBe(2);

    await expect(
      service.execute('find and replace GST in column D', 'FIND_REPLACE', workbookContext),
    ).rejects.toThrow('numeric_find_replace_escalation_required');

    expect(openRouter.complete).not.toHaveBeenCalled();
  });

  it('allows non-financial find-replace', async () => {
    openRouter.complete.mockResolvedValue(
      JSON.stringify({
        answer: 'Replaced ABC with XYZ.',
        actions: [
          {
            type: 'BATCH_SET',
            sheetName: 'Sheet1',
            operations: [{ address: 'B2', value: 'XYZ' }],
          },
        ],
      }),
    );

    const result = await service.execute(
      'find and replace ABC with XYZ',
      'FIND_REPLACE',
      workbookContext,
    );

    expect(openRouter.complete).toHaveBeenCalledTimes(1);
    expect(result.actions[0].type).toBe('BATCH_SET');
  });
});
