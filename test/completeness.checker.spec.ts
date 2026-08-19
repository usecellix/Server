import { CompletenessChecker } from '../src/agents/checkers/completeness.checker';
import { SubTask } from '../src/agents/types/agent.types';

describe('CompletenessChecker', () => {
  const checker = new CompletenessChecker();

  const subtasks: SubTask[] = [
    {
      id: 's1',
      description: 'Add row',
      targetSheet: 'Sheet1',
      dependsOn: [],
      estimatedActions: 1,
    },
    {
      id: 's2',
      description: 'Format header',
      targetSheet: 'Sheet1',
      dependsOn: ['s1'],
      estimatedActions: 2,
    },
  ];

  it('fails when a subtask has no actions', () => {
    const result = checker.check(subtasks, [
      { subtask: subtasks[0], actions: [{ type: 'ADD_ROW', data: ['X'] }] },
      { subtask: subtasks[1], actions: [] },
    ]);

    expect(result.passed).toBe(false);
    expect(result.subtaskResults[1].passed).toBe(false);
  });

  it('passes when every subtask has actions', () => {
    const result = checker.check(subtasks, [
      { subtask: subtasks[0], actions: [{ type: 'ADD_ROW', data: ['X'] }] },
      { subtask: subtasks[1], actions: [{ type: 'FORMAT_RANGE', row: 0, col: 0 }] },
    ]);

    expect(result.passed).toBe(true);
  });

  it('fails a subtask whose Executor output had actions discarded as unusable', () => {
    const result = checker.check(subtasks, [
      { subtask: subtasks[0], actions: [{ type: 'ADD_ROW', data: ['X'] }] },
      {
        subtask: subtasks[1],
        actions: [{ type: 'FORMAT_RANGE', row: 0, col: 0 }],
        droppedActions: [{ rawType: 'APPLY_AUTOFILTER', reason: 'unknown-type' }],
      },
    ]);

    // Actions were present, so the old count-based checks pass — only the
    // dropped-action check catches this partial completion.
    expect(result.passed).toBe(false);
    expect(result.subtaskResults[1].passed).toBe(false);
    expect(result.subtaskResults[1].feedback).toContain('APPLY_AUTOFILTER');
  });

  it('tells the retry exactly what a missing-required-fields drop needs, not the generic unknown-type advice', () => {
    const result = checker.check(subtasks, [
      { subtask: subtasks[0], actions: [{ type: 'ADD_ROW', data: ['X'] }] },
      {
        subtask: subtasks[1],
        actions: [],
        droppedActions: [{ rawType: 'BATCH_SET', reason: 'missing-required-fields' }],
      },
    ]);

    expect(result.passed).toBe(false);
    const issue = result.subtaskResults[1].issues[0];
    expect(issue.description).toContain('missing a required field');
    expect(issue.suggestion).toContain('operations');
    expect(issue.suggestion).not.toContain('Available action types list');
  });

  it('combines both drop reasons into one message when a subtask has both kinds', () => {
    const result = checker.check(subtasks, [
      { subtask: subtasks[0], actions: [{ type: 'ADD_ROW', data: ['X'] }] },
      {
        subtask: subtasks[1],
        actions: [{ type: 'FORMAT_RANGE', row: 0, col: 0 }],
        droppedActions: [
          { rawType: 'APPLY_AUTOFILTER', reason: 'unknown-type' },
          { rawType: 'BATCH_SET', reason: 'missing-required-fields' },
        ],
      },
    ]);

    const issue = result.subtaskResults[1].issues[0];
    expect(issue.description).toContain('APPLY_AUTOFILTER');
    expect(issue.description).toContain('missing a required field');
    expect(issue.suggestion).toContain('Available action types list');
    expect(issue.suggestion).toContain('operations');
  });
});
