import { OrchestratorService } from '../src/agents/orchestrator.service';
import { PlannerAgent } from '../src/agents/planner.agent';
import { AgenticLoopService } from '../src/agents/agenticLoop.service';
import { SpecExtractorAgent } from '../src/agents/spec-extractor.agent';
import { SseEmitter } from '../src/agents/sse.emitter';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * Live incident (TASKS.md #281): a turn resuming after a clarifying question
 * sends only the short reply ("dd-mm-yyyy") as its prompt — the original long
 * request lives in `conversationHistory`. `withBuildSpec` used to pass ONLY
 * the current turn's short prompt to `SpecExtractorAgent`, so its length/
 * subtask-count gate saw a 10-character message and Phase 1/1.5's reliability
 * protections silently never engaged for a resumed long build — exactly the
 * turn most likely to need them, since it is a continuation of the same big
 * build that already needed a clarifying question asked.
 */

const mockContext: WorkbookContext = {
  activeSheetName: 'Sheet1',
  sheets: [{
    name: 'Sheet1', usedRange: 'A1:A1', rowCount: 1, columnCount: 1,
    values: [['']], formulas: [['']], numberFormats: [['General']],
    structure: 'data_table', headerRowIndex: 0,
  }],
  namedRanges: [],
  tables: [],
};

const LONG_ORIGINAL_PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the details of the ' +
  'remaining sheets, in the main sheet i need to have dashboard also, my need to record payments and related things ' +
  ',which all month sheets include Unit No, Guest, Guest name, check in, check out, Rate per night, total amount, ' +
  'source, payment status, bank account';

describe('OrchestratorService — build-spec sees the real conversation, not just the reply (TASKS.md #281)', () => {
  const plan = {
    subtasks: [
      { id: 's1', description: 'Create sheet Main', targetSheet: 'Main', dependsOn: [], estimatedActions: 3 },
      { id: 's2', description: 'Create sheet Lists', targetSheet: 'Lists', dependsOn: [], estimatedActions: 3 },
      { id: 's3', description: "Create sheet 'January' with headers", targetSheet: 'January', dependsOn: [], estimatedActions: 3 },
    ],
    clarificationsNeeded: [],
    confidence: 'high' as const,
    reasoning: '',
  };

  it('planForStepwiseRun passes the reconstructed prompt (history + reply), not the short reply alone', async () => {
    const planner = { plan: jest.fn().mockResolvedValue(plan) };
    const agenticLoop = {};
    const specExtractor = { attach: jest.fn().mockResolvedValue(plan) };

    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
      specExtractor as unknown as SpecExtractorAgent,
    );

    await orchestrator.planForStepwiseRun(
      {
        prompt: 'dd-mm-yyyy',
        context: mockContext,
        conversationHistory: [
          { role: 'user', content: LONG_ORIGINAL_PROMPT },
          { role: 'assistant', content: 'What date format do you want?' },
        ],
      },
      new SseEmitter(() => {}),
    );

    expect(specExtractor.attach).toHaveBeenCalledTimes(1);
    const [effectivePrompt] = specExtractor.attach.mock.calls[0];
    expect(effectivePrompt).toContain(LONG_ORIGINAL_PROMPT);
    expect(effectivePrompt).toContain('dd-mm-yyyy');
    expect(effectivePrompt.length).toBeGreaterThanOrEqual(200); // clears shouldExtractBuildSpec's own gate
  });

  it('a fresh request with no history still gets just its own prompt (no behavior change for the common case)', async () => {
    const planner = { plan: jest.fn().mockResolvedValue(plan) };
    const agenticLoop = {};
    const specExtractor = { attach: jest.fn().mockResolvedValue(plan) };

    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
      specExtractor as unknown as SpecExtractorAgent,
    );

    await orchestrator.planForStepwiseRun(
      { prompt: LONG_ORIGINAL_PROMPT, context: mockContext },
      new SseEmitter(() => {}),
    );

    expect(specExtractor.attach.mock.calls[0][0]).toBe(LONG_ORIGINAL_PROMPT);
  });

  it('only USER turns are folded in — an assistant clarification question never pollutes the grounding text', async () => {
    const planner = { plan: jest.fn().mockResolvedValue(plan) };
    const agenticLoop = {};
    const specExtractor = { attach: jest.fn().mockResolvedValue(plan) };

    const orchestrator = new OrchestratorService(
      planner as unknown as PlannerAgent,
      agenticLoop as unknown as AgenticLoopService,
      specExtractor as unknown as SpecExtractorAgent,
    );

    await orchestrator.planForStepwiseRun(
      {
        prompt: 'dd-mm-yyyy',
        context: mockContext,
        conversationHistory: [
          { role: 'assistant', content: 'What date format do you want?' },
          { role: 'user', content: LONG_ORIGINAL_PROMPT },
        ],
      },
      new SseEmitter(() => {}),
    );

    const [effectivePrompt] = specExtractor.attach.mock.calls[0];
    expect(effectivePrompt).not.toContain('What date format');
  });
});
