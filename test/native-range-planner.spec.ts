import * as fs from 'fs';
import * as path from 'path';
import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { buildExecutorUserMessage } from '../src/agents/prompts/executor.prompt';
import { PlannerOutput, SubTask, WorkbookContext } from '../src/agents/types/agent.types';

const fixtures = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'fixtures', 'native-range-planner.json'), 'utf8'),
) as Array<{
  prompt: string;
  expectedSubtaskCount?: number;
  expectedSubtaskCountMax?: number;
  expectedSuggestedActionType: string;
  examplePlan: PlannerOutput;
}>;

describe('native range planner contract', () => {
  const agent = new PlannerAgent({} as OpenRouterService, {} as AppConfigService);

  it.each(fixtures)('preserves suggestedActionType for: $prompt', (fixture) => {
    const normalized = agent.normalizePlannerOutputForTest(fixture.examplePlan);
    const maxCount = fixture.expectedSubtaskCountMax ?? fixture.expectedSubtaskCount ?? 2;
    expect(normalized.subtasks.length).toBeLessThanOrEqual(maxCount);
    if (fixture.expectedSubtaskCount !== undefined) {
      expect(normalized.subtasks.length).toBe(fixture.expectedSubtaskCount);
    }

    const withSuggestion = normalized.subtasks.find(
      (s) => s.suggestedActionType === fixture.expectedSuggestedActionType,
    );
    expect(withSuggestion).toBeDefined();
    expect(withSuggestion?.estimatedActions).toBe(1);
    expect(normalized.subtasks.length).toBeLessThanOrEqual(2);
  });

  it('surfaces suggestedActionType in the executor user message', () => {
    const subtask: SubTask = {
      id: 's2',
      description: 'Copy pending rows',
      targetSheet: 'Pending Payments',
      dependsOn: ['s1'],
      estimatedActions: 1,
      suggestedActionType: 'COPY_FILTERED_RANGE',
    };
    const context: WorkbookContext = {
      activeSheetName: 'Purchase Register',
      sheets: [
        {
          name: 'Pending Payments',
          usedRange: 'A1',
          rowCount: 1,
          columnCount: 1,
          values: [['']],
          formulas: [['']],
          numberFormats: [['General']],
          structure: 'unknown',
        },
      ],
      namedRanges: [],
      tables: [],
      onDemandFetchEnabled: true,
    };

    const message = buildExecutorUserMessage(subtask, context, []);
    expect(message).toContain('Suggested action type: COPY_FILTERED_RANGE');
    expect(message).toContain('do not use get_range_data');
  });
});
