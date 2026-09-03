import { PlannerAgent } from '../src/agents/planner.agent';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Confirms the reasoning.effort split is wired where it actually matters —
 * each agent's own complete() call args, not just the OpenRouterService
 * default. Planner stays at full capability (explicit instruction — do not
 * lower it to save tokens); Executor is explicit 'low' rather than relying on
 * OpenRouterService's own 'low' default, since a caller-level default is easy
 * to silently change later without anyone noticing this agent depended on it.
 */

const context: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [],
  namedRanges: [],
  tables: [],
};

describe('PlannerAgent — reasoning effort stays at full capability', () => {
  function buildAgent(completeImpl: jest.Mock): PlannerAgent {
    const llm = { complete: completeImpl } as unknown as OpenRouterService;
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    return new PlannerAgent(llm, config);
  }

  it('requests reasoningEffort "high" on the first planning call', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({
        subtasks: [
          { id: 's1', description: 'Add row', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 1 },
        ],
        clarificationsNeeded: [],
        confidence: 'high',
        reasoning: 'ok',
      }),
    );
    const agent = buildAgent(complete);

    await agent.plan('Add a row', context, [], undefined, 'corr_1');

    expect(complete).toHaveBeenCalledTimes(1);
    const firstCall = complete.mock.calls[0][0] as { reasoningEffort?: string };
    expect(firstCall.reasoningEffort).toBe('high');
  });

  it('keeps reasoningEffort "high" on the parse-failure retry, not a downgraded value', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce('not valid json')
      .mockResolvedValueOnce(
        JSON.stringify({
          subtasks: [
            { id: 's1', description: 'Add row', targetSheet: 'Sheet1', dependsOn: [], estimatedActions: 1 },
          ],
          clarificationsNeeded: [],
          confidence: 'high',
          reasoning: 'ok',
        }),
      );
    const agent = buildAgent(complete);

    await agent.plan('Add a row', context, [], undefined, 'corr_2');

    expect(complete).toHaveBeenCalledTimes(2);
    const retryCall = complete.mock.calls[1][0] as { reasoningEffort?: string };
    expect(retryCall.reasoningEffort).toBe('high');
  });
});

describe('ExecutorAgent — reasoning effort explicitly low', () => {
  function buildAgent(completeImpl: jest.Mock): ExecutorAgent {
    const llm = { complete: completeImpl } as unknown as OpenRouterService;
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    return new ExecutorAgent(llm, config);
  }

  const subtask: SubTask = {
    id: 's1',
    description: 'Add row',
    targetSheet: 'Sheet1',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('requests reasoningEffort "low" explicitly on the first executor call', async () => {
    const complete = jest.fn().mockResolvedValue(
      JSON.stringify({ subtaskId: 's1', actions: [{ type: 'ADD_ROW', data: ['X'] }], isDone: true }),
    );
    const agent = buildAgent(complete);

    await agent.execute(subtask, context, [], 'corr_3');

    expect(complete).toHaveBeenCalledTimes(1);
    const firstCall = complete.mock.calls[0][0] as { reasoningEffort?: string };
    expect(firstCall.reasoningEffort).toBe('low');
  });

  it('keeps reasoningEffort "low" on the JSON parse-failure retry', async () => {
    const complete = jest
      .fn()
      .mockResolvedValueOnce('not valid json')
      .mockResolvedValueOnce(
        JSON.stringify({ subtaskId: 's1', actions: [{ type: 'ADD_ROW', data: ['X'] }], isDone: true }),
      );
    const agent = buildAgent(complete);

    await agent.execute(subtask, context, [], 'corr_4');

    expect(complete).toHaveBeenCalledTimes(2);
    const retryCall = complete.mock.calls[1][0] as { reasoningEffort?: string };
    expect(retryCall.reasoningEffort).toBe('low');
  });
});
