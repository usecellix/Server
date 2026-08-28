import {
  PlannerAgent,
  buildMultiClauseClarification,
  splitWriteClauses,
} from '../src/agents/planner.agent';
import { PlannerOutput } from '../src/agents/types/agent.types';

/**
 * Task #91 (2026-08-27): splitting on a bare "and" treated a conjunction joining two
 * objects of one question as two instructions, and the clarification it produced was
 * hardcoded to an unrelated column-deletion scenario.
 */
describe('splitWriteClauses — conjunctions of nouns are not separate commands (#91)', () => {
  const REAL_PROMPT =
    "What's in this sheet? Give me total invoice value, total tax collected, " +
    'how many invoices are pending payment and what they add up to.';

  it('does not split the reported read-only prompt', () => {
    expect(splitWriteClauses(REAL_PROMPT)).toEqual([]);
  });

  it.each([
    'Show me the pending invoices and what they add up to',
    'How many rows are Paid and how many are Pending?',
    'What is the total and is it correct?',
  ])('treats %j as a single clause', (prompt) => {
    expect(splitWriteClauses(prompt)).toEqual([]);
  });
});

describe('splitWriteClauses — genuine multi-command prompts still split (#91)', () => {
  it('still splits the Spec 22 delete+annotate compound', () => {
    const clauses = splitWriteClauses(
      'delete the column payment status and in remarks add priority to unpaid invoices',
    );
    expect(clauses.length).toBeGreaterThanOrEqual(2);
    expect(clauses[0]).toMatch(/delete/i);
    expect(clauses[1]).toMatch(/remarks|priority/i);
  });

  it.each([
    'sort by supplier name and highlight the pending rows in red',
    'create a summary sheet then add a total row to it',
    'clear the remarks column and also format the header bold',
  ])('splits %j into commands', (prompt) => {
    expect(splitWriteClauses(prompt).length).toBeGreaterThanOrEqual(2);
  });
});

describe('buildMultiClauseClarification — no invented operations (#91)', () => {
  it('never mentions column deletion for a prompt that did not ask for it', () => {
    const clauses = [
      'fill the remarks column for pending invoices',
      'highlight them in red',
    ];
    const text = buildMultiClauseClarification(clauses, [clauses[1]], 1);
    expect(text).not.toMatch(/column deletion|annotate\/filter/i);
    expect(text).toMatch(/highlight them in red/);
  });

  it('names the clauses it could not plan', () => {
    const clauses = ['sort by supplier', 'add a total row'];
    const text = buildMultiClauseClarification(clauses, ['add a total row'], 1);
    expect(text).toContain('add a total row');
  });

  it('agrees in number with the planned step count', () => {
    expect(buildMultiClauseClarification(['a b c d', 'e f g h'], [], 1)).toMatch(/1 step\b/);
    expect(buildMultiClauseClarification(['a b c d', 'e f g h'], [], 2)).toMatch(/2 steps\b/);
  });
});

describe('ensureMultiClauseCoverage — read-only prompts are not flagged (#91)', () => {
  const agent = Object.create(PlannerAgent.prototype) as PlannerAgent;

  it('leaves a single-subtask plan alone for a question prompt', () => {
    const plan: PlannerOutput = {
      subtasks: [
        {
          id: 's1',
          description: 'Report totals for the Purchase Register',
          targetSheet: 'Purchase Register',
          dependsOn: [],
          estimatedActions: 1,
          suggestedActionType: 'SET_FORMULA',
        },
      ],
      clarificationsNeeded: [],
      confidence: 'high',
      reasoning: 'Answer the question',
    };

    const result = agent.ensureMultiClauseCoverage(
      'how many invoices are pending payment and what they add up to',
      plan,
    );

    expect(result.clarificationsNeeded).toHaveLength(0);
    expect(result.confidence).toBe('high');
  });
});
