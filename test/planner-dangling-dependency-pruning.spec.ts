import { PlannerAgent } from '../src/agents/planner.agent';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';
import { AppConfigService } from '../src/config/app-config.service';
import { PlannerOutput, SubTask } from '../src/agents/types/agent.types';

/**
 * Live incident (Sept 8, 2026): a 19-subtask plan's raw response ended
 * mid-generation while describing subtask s14 ("July") — the model's own
 * closing brackets made the truncated JSON syntactically VALID, so the
 * existing truncation-retry ladder (which only trusts the provider's
 * `finishReason`) never fired. 14 subtasks came back, but s2's `dependsOn`
 * still named `s15`..`s19` (August through December) — subtasks that were
 * never generated at all. Nothing caught the mismatch: `normalizePlannerOutput`
 * only checks each subtask individually has id/description/targetSheet, never
 * that `dependsOn` edges resolve to real subtasks.
 *
 * The consequence reached the user's real workbook: `computeExecutionWaves`'s
 * "stranded" fallback ran the unsatisfiable subtasks anyway (nothing else
 * could ever become newly eligible), and their formulas — `=SUM(August!G:G)`
 * etc. — landed as 18 `#REF!` errors on sheets that were never created.
 *
 * `pruneUnsatisfiableSubtasks` is the fix: after parsing, drop any subtask
 * whose dependsOn (directly or transitively) can never be satisfied, and
 * report what was dropped via `clarificationsNeeded` rather than silently
 * shipping broken formulas as if the build succeeded.
 */

const agent = new PlannerAgent({} as OpenRouterService, {} as AppConfigService);

function subtask(id: string, dependsOn: string[] = [], targetSheet = 'Main'): SubTask {
  return { id, description: `subtask ${id}`, targetSheet, dependsOn, estimatedActions: 1 };
}

function plan(subtasks: SubTask[]): PlannerOutput {
  return { subtasks, clarificationsNeeded: [], confidence: 'high', reasoning: '' };
}

describe('PlannerAgent.pruneUnsatisfiableSubtasks', () => {
  it('leaves a plan with no dangling dependencies completely untouched', () => {
    const input = plan([subtask('s1'), subtask('s2', ['s1'])]);
    const result = agent.pruneUnsatisfiableSubtasksForTest(input);
    expect(result).toBe(input); // same reference — genuinely a no-op, not just equal
  });

  it('reproduces the exact live incident: drops subtasks referencing never-generated month sheets', () => {
    // s8-s14 = January..July (generated). s15-s19 (August..December) never
    // appear in `subtasks` at all, but s2 and s7 still list them in dependsOn —
    // exactly the shape the truncated response produced.
    const input = plan([
      subtask('s1', ['s8', 's9', 's10', 's11', 's12', 's13']),
      subtask('s2', ['s1', 's14', 's15', 's16', 's17', 's18', 's19']),
      subtask('s3', ['s2']),
      subtask('s4', ['s3']),
      subtask('s5', ['s4']),
      subtask('s6', ['s5']),
      subtask('s7', ['s8', 's9', 's10', 's11', 's12', 's13', 's14', 's15', 's16', 's17', 's18', 's19'], 'Lists'),
      subtask('s8', [], 'January'),
      subtask('s9', [], 'February'),
      subtask('s10', [], 'March'),
      subtask('s11', [], 'April'),
      subtask('s12', [], 'May'),
      subtask('s13', [], 'June'),
      subtask('s14', [], 'July'),
    ]);

    const result = agent.pruneUnsatisfiableSubtasksForTest(input);

    // s2 depends directly on missing s15-s19 -> pruned. s3-s7 depend
    // (transitively, via s2) on something unsatisfiable -> also pruned.
    // s1 and the 7 month-create subtasks (s8-s14) have no dangling refs at
    // all -> kept, exactly matching what the live incident's wave 0 delivered
    // successfully.
    const keptIds = result.subtasks.map((s) => s.id).sort();
    expect(keptIds).toEqual(['s1', 's10', 's11', 's12', 's13', 's14', 's8', 's9'].sort());
  });

  it('reports what was dropped via clarificationsNeeded, not silently', () => {
    const input = plan([subtask('s1', ['s99'])]);
    const result = agent.pruneUnsatisfiableSubtasksForTest(input);

    expect(result.subtasks).toEqual([]);
    expect(result.clarificationsNeeded).toHaveLength(1);
    expect(result.clarificationsNeeded[0]).toMatch(/could not fully plan/i);
  });

  it('prunes a subtask depending on ANOTHER pruned subtask, not just the one with the direct dangling ref', () => {
    // s3 -> s2 -> s99 (missing). s3's OWN dependsOn ids all exist (just s2),
    // but s2 itself can never complete, so s3 must be pruned too.
    const input = plan([subtask('s2', ['s99']), subtask('s3', ['s2'])]);
    const result = agent.pruneUnsatisfiableSubtasksForTest(input);
    expect(result.subtasks).toEqual([]);
  });

  it('keeps independent subtasks that do not depend on the unsatisfiable branch', () => {
    const input = plan([
      subtask('s1', ['s99']), // unsatisfiable
      subtask('independent', []), // untouched
    ]);
    const result = agent.pruneUnsatisfiableSubtasksForTest(input);
    expect(result.subtasks.map((s) => s.id)).toEqual(['independent']);
  });
});
