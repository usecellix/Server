import { computeExecutionWaves } from '../src/agents/utils/task-graph.util';
import { SubTask } from '../src/agents/types/agent.types';

describe('computeExecutionWaves', () => {
  it('groups independent subtasks into the same wave', () => {
    const subtasks: SubTask[] = [
      {
        id: 's1',
        description: 'Format Summary',
        targetSheet: 'Summary',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 's2',
        description: 'Format Data',
        targetSheet: 'Data',
        dependsOn: [],
        estimatedActions: 1,
      },
      {
        id: 's3',
        description: 'Update totals',
        targetSheet: 'Summary',
        dependsOn: ['s1'],
        estimatedActions: 1,
      },
    ];

    const waves = computeExecutionWaves(subtasks);

    expect(waves).toHaveLength(2);
    expect(waves[0].map((subtask) => subtask.id).sort()).toEqual(['s1', 's2']);
    expect(waves[1].map((subtask) => subtask.id)).toEqual(['s3']);
  });
});
