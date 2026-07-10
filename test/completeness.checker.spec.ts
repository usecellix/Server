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
});
