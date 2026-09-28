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
      headerRowIndex: 0,
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
      // Usage-accumulator out-param — populates conversation.service.ts's
      // audit-log telemetry with real Planner/Executor/Verifier token usage.
      expect.objectContaining({ promptTokens: 0, completionTokens: 0, totalTokens: 0 }),
      // Short-progress callback (TASKS.md #193) — only consulted by two-pass
      // planning, but always passed through so that path can use it.
      expect.any(Function),
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

  /** Builds a plan with `count` trivial subtasks, all pointing at Sheet1. */
  function buildPlan(count: number, clarificationsNeeded: string[]) {
    return {
      subtasks: Array.from({ length: count }, (_, i) => ({
        id: `s${i + 1}`,
        description: `Step ${i + 1}`,
        targetSheet: 'Sheet1',
        dependsOn: [] as string[],
        estimatedActions: 1,
      })),
      clarificationsNeeded,
      confidence: 'high' as const,
      reasoning: 'ok',
    };
  }

  // TASKS.md #259 — a narrow addition on top of #171: a BIG build (>= 6
  // subtasks, matching the live 12-month-sheet-ledger case that prompted
  // this) with real open questions now blocks and asks first, instead of
  // guessing (e.g. "Unassigned" bank names) across ~150 actions and hoping
  // the summary note gets read.
  it('blocks a big build with real open questions (TASKS.md #259)', async () => {
    const planner = {
      plan: jest.fn().mockResolvedValue(
        buildPlan(6, ['What are the real bank account names for the dropdown?']),
      ),
    };
    const agenticLoop = { run: jest.fn() };
    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
    );

    const sseEvents: Array<{ event: string; data: Record<string, unknown> }> = [];
    const result = await orchestrator.runDetailed(
      { prompt: 'build 12 month sheets plus a dashboard', context: mockContext },
      new SseEmitter((event, data) => sseEvents.push({ event, data })),
    );

    expect(result.clarificationRequested).toBe(true);
    expect(result.actions).toEqual([]);
    expect(agenticLoop.run).not.toHaveBeenCalled();
    expect(result.openQuestions).toEqual([
      'What are the real bank account names for the dropdown?',
    ]);
    expect(sseEvents.some((e) => e.event === 'clarification')).toBe(true);
  });

  // #171's original case: a SMALL plan with open questions must still
  // proceed under an assumption rather than block — this is the exact
  // behavior #259 is deliberately scoped to leave untouched.
  it('does NOT block a small plan with open questions (preserves TASKS.md #171)', async () => {
    const expectedActions: Action[] = [
      { type: 'SET_CELL', row: 1, col: 1, value: 1200, sheetName: 'Sheet1' } as Action,
    ];
    const planner = {
      plan: jest.fn().mockResolvedValue(
        buildPlan(2, ['Which year should this use?']),
      ),
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
    const result = await orchestrator.runDetailed(
      { prompt: 'do a couple small things', context: mockContext },
      new SseEmitter((event, data) => sseEvents.push({ event, data })),
    );

    expect(result.clarificationRequested).toBe(false);
    expect(agenticLoop.run).toHaveBeenCalledTimes(1);
    expect(result.actions).toEqual(expectedActions);
    expect(sseEvents.some((e) => e.event === 'clarification')).toBe(false);
  });
});
