import { ExecutorAgent } from '../src/agents/executor.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';
import { resolveExecutorMaxTokens } from '../src/agents/utils/executor-token-budget.util';

/**
 * Confirms ExecutorAgent actually wires subtask.estimatedActions into the LLM
 * call's maxTokens, rather than the fix living only in the (correctly tested,
 * but otherwise unused) pure util function.
 */
describe('ExecutorAgent — scales maxTokens with subtask.estimatedActions', () => {
  const context: WorkbookContext = {
    activeSheetName: 'Main',
    sheets: [],
    namedRanges: [],
    tables: [],
    onDemandFetchEnabled: false,
  };

  function buildAgent() {
    const complete = jest.fn(async (_opts: { maxTokens?: number }) =>
      JSON.stringify({ subtaskId: 's1', actions: [], isDone: true }),
    );
    const openRouter = { complete } as unknown as OpenRouterService;
    const config = { openRouterModelHigh: 'openai/gpt-5' } as unknown as AppConfigService;
    const agent = new ExecutorAgent(openRouter, config);
    return { agent, complete };
  }

  it('requests a larger budget for a subtask estimating many actions', async () => {
    const { agent, complete } = buildAgent();
    const subtask: SubTask = {
      id: 's1',
      description: 'Write 18 formula cells for the Jan-Jun Monthly Totals table',
      targetSheet: 'Main',
      dependsOn: [],
      estimatedActions: 18,
    };

    await agent.execute(subtask, context);

    expect(complete).toHaveBeenCalledTimes(1);
    expect(complete.mock.calls[0][0]).toMatchObject({
      maxTokens: resolveExecutorMaxTokens(18),
    });
    expect(complete.mock.calls[0][0].maxTokens).toBeGreaterThan(2000);
  });

  it('uses the base budget for a small subtask', async () => {
    const { agent, complete } = buildAgent();
    const subtask: SubTask = {
      id: 's1',
      description: 'Set A1 to Total',
      targetSheet: 'Main',
      dependsOn: [],
      estimatedActions: 1,
    };

    await agent.execute(subtask, context);

    expect(complete.mock.calls[0][0]).toMatchObject({
      maxTokens: resolveExecutorMaxTokens(1),
    });
  });
});
