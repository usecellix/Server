import {
  PlannerAgent,
  clauseLikelyCovered,
  splitWriteClauses,
} from '../src/agents/planner.agent';
import { PlannerOutput } from '../src/agents/types/agent.types';

describe('Spec 22 planner multi-clause coverage', () => {
  const agent = Object.create(PlannerAgent.prototype) as PlannerAgent;

  it('splitWriteClauses breaks delete+annotate compounds', () => {
    const clauses = splitWriteClauses(
      'delete the column payment status and in remarks add priority to unpaid invoices',
    );
    expect(clauses.length).toBeGreaterThanOrEqual(2);
    expect(clauses[0]).toMatch(/delete/i);
    expect(clauses[1]).toMatch(/remarks|priority/i);
  });

  it('flags incomplete single-subtask plans for multi-clause prompts', () => {
    const plan: PlannerOutput = {
      subtasks: [
        {
          id: 's1',
          description: 'Delete the Payment Status column',
          targetSheet: 'Purchase Register',
          dependsOn: [],
          estimatedActions: 1,
          suggestedActionType: 'DELETE_COLUMN',
        },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Only planned delete',
    };

    const covered = agent.ensureMultiClauseCoverage(
      'delete the column payment status and in remarks add priority to unpaid invoices',
      plan,
    );

    expect(covered.clarificationsNeeded.length).toBeGreaterThan(0);
    expect(covered.confidence).toBe('low');
  });

  it('wires DELETE_COLUMN dependsOn annotate when both subtasks exist', () => {
    const plan: PlannerOutput = {
      subtasks: [
        {
          id: 's1',
          description: 'Delete the Payment Status column',
          targetSheet: 'Purchase Register',
          dependsOn: [],
          estimatedActions: 1,
          suggestedActionType: 'DELETE_COLUMN',
        },
        {
          id: 's2',
          description: 'Set Remarks to Priority where Payment Status is unpaid',
          targetSheet: 'Purchase Register',
          dependsOn: [],
          estimatedActions: 1,
          suggestedActionType: 'SET_MATCHING_ROWS',
        },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Both clauses',
    };

    const covered = agent.ensureMultiClauseCoverage(
      'delete the column payment status and in remarks add priority to unpaid invoices',
      plan,
    );

    const del = covered.subtasks.find((s) => s.suggestedActionType === 'DELETE_COLUMN');
    const annotate = covered.subtasks.find((s) => s.suggestedActionType === 'SET_MATCHING_ROWS');
    expect(del?.dependsOn).toContain(annotate!.id);
    expect(covered.clarificationsNeeded).toEqual([]);
  });

  it('clauseLikelyCovered matches shared tokens', () => {
    expect(
      clauseLikelyCovered(
        'in remarks add priority to unpaid invoices',
        'Set Remarks to Priority where Payment Status is unpaid',
      ),
    ).toBe(true);
  });
});
