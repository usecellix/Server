import { OrchestratorService } from '../src/agents/orchestrator.service';
import { PlannerAgent } from '../src/agents/planner.agent';
import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { SseEmitter } from '../src/agents/sse.emitter';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';

const mockContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      usedRange: 'A1:B2',
      rowCount: 2,
      columnCount: 2,
      values: [
        ['Item', 'Amount'],
        ['Sales', 1000],
      ],
      formulas: [['', ''], ['', '']],
      numberFormats: [['General', 'General']],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('OrchestratorService SSE write path', () => {
  it('streams planned actions through planner → agentic loop → SSE events', async () => {
    const expectedActions: Action[] = [
      { type: 'SET_CELL', row: 1, col: 1, value: 1200, sheetName: 'Sheet1' } as Action,
    ];

    const planner = {
      plan: jest.fn().mockResolvedValue({
        subtasks: [
          {
            id: 's1',
            description: 'Update sales amount',
            targetSheet: 'Sheet1',
            dependsOn: [],
            estimatedActions: 1,
          },
        ],
        clarificationsNeeded: [],
        confidence: 'high',
        reasoning: 'Single cell update',
      }),
    };

    const agenticLoop = {
      run: jest.fn().mockResolvedValue({
        actions: expectedActions,
        iterationsRun: 1,
        verifierPassed: true,
        completedSubtasks: [],
        failedSubtask: null,
        partialProgress: false,
      }),
    };

    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
    );

    const sseEvents: Array<{ event: string; data: Record<string, unknown> }> = [];
    const emitter = new SseEmitter((event, data) => {
      sseEvents.push({ event, data });
    });

    const actions = await orchestrator.run(
      {
        prompt: 'Set sales to 1200',
        context: mockContext,
        conversationHistory: [],
      },
      emitter,
    );

    expect(planner.plan).toHaveBeenCalledWith(
      'Set sales to 1200',
      mockContext,
      [],
      undefined,
      expect.any(String),
      undefined,
      undefined,
    );
    expect(agenticLoop.run).toHaveBeenCalledTimes(1);
    expect(actions).toEqual(expectedActions);

    const eventNames = sseEvents.map((e) => e.event);
    expect(eventNames).toContain('thinking');
    expect(eventNames).toContain('status');
  });

  it('emits clarification SSE when planner needs more info', async () => {
    const planner = {
      plan: jest.fn().mockResolvedValue({
        subtasks: [],
        clarificationsNeeded: ['Which column should be updated?'],
        confidence: 'low',
        reasoning: 'Ambiguous',
      }),
    };

    const agenticLoop = { run: jest.fn() };

    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
    );

    const sseEvents: Array<{ event: string; data: Record<string, unknown> }> = [];
    const actions = await orchestrator.run(
      { prompt: 'Update the value', context: mockContext },
      new SseEmitter((event, data) => sseEvents.push({ event, data })),
    );

    expect(actions).toEqual([]);
    expect(agenticLoop.run).not.toHaveBeenCalled();
    expect(sseEvents.some((e) => e.event === 'clarification')).toBe(true);
  });

  it('blocks Executor when confidence is low even without clarificationsNeeded', async () => {
    const planner = {
      plan: jest.fn().mockResolvedValue({
        subtasks: [
          {
            id: 's1',
            description: 'Guess clean-up',
            targetSheet: 'Sheet1',
            dependsOn: [],
            estimatedActions: 3,
          },
        ],
        clarificationsNeeded: [],
        confidence: 'low',
        reasoning: "Unclear what 'clean up' means",
      }),
    };

    const agenticLoop = { run: jest.fn() };
    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
    );

    const sseEvents: Array<{ event: string; data: Record<string, unknown> }> = [];
    const result = await orchestrator.runDetailed(
      { prompt: 'clean up this data', context: mockContext },
      new SseEmitter((event, data) => sseEvents.push({ event, data })),
    );

    expect(result.clarificationRequested).toBe(true);
    expect(result.actions).toEqual([]);
    expect(agenticLoop.run).not.toHaveBeenCalled();
    expect(sseEvents.some((e) => e.event === 'clarification')).toBe(true);
  });
});
