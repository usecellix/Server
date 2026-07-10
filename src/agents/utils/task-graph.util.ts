import { SubTask } from '../types/agent.types';

/** Group subtasks into dependency waves — each wave can run in parallel. */
export function computeExecutionWaves(subtasks: SubTask[]): SubTask[][] {
  const byId = new Map(subtasks.map((subtask) => [subtask.id, subtask]));
  const completed = new Set<string>();
  const remaining = new Set(subtasks.map((subtask) => subtask.id));
  const waves: SubTask[][] = [];

  while (remaining.size > 0) {
    const wave = subtasks.filter(
      (subtask) =>
        remaining.has(subtask.id) &&
        subtask.dependsOn.every((dep) => completed.has(dep) && byId.has(dep)),
    );

    if (wave.length === 0) {
      break;
    }

    waves.push(wave);
    for (const subtask of wave) {
      remaining.delete(subtask.id);
      completed.add(subtask.id);
    }
  }

  if (remaining.size > 0) {
    const stranded = subtasks.filter((subtask) => remaining.has(subtask.id));
    waves.push(stranded);
  }

  return waves;
}

export function collectDependencyIds(subtasks: SubTask[]): Set<string> {
  const ids = new Set<string>();
  for (const subtask of subtasks) {
    ids.add(subtask.id);
    for (const dep of subtask.dependsOn) {
      ids.add(dep);
    }
  }
  return ids;
}
