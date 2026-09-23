import {
  cloneActionsForSheet,
  findCloneGroups,
} from '../src/agents/utils/template-replicate.util';
import { Action, SubTask } from '../src/agents/types/agent.types';

// Phase 2 of LONG_PROMPT_RELIABILITY_PLAN.md — shapes copied from the live plan
// (agent_runs run_1789970471833_2i8cy2g): 12 month subtasks identical except
// for the sheet name and the 3-letter table abbreviation.
const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const monthSubtask = (month: string, i: number, extra = ''): SubTask => ({
  id: `p2_s${i + 1}`,
  targetSheet: month,
  dependsOn: ['p1_s1', 'p1_s2'],
  estimatedActions: 22,
  description:
    `Create sheet '${month}' and build the booking/payment log: headers in row 1 — Unit No | Guest. ` +
    `Create Excel Table 'tbl${month.slice(0, 3)}' over A1:M2. Data validation lists on Source (Lists!$B$3:$B$20).${extra}`,
});

describe('findCloneGroups', () => {
  it('groups the 12 live month subtasks into one template + 11 clones', () => {
    const groups = findCloneGroups(MONTHS.map((m, i) => monthSubtask(m, i)));
    expect(groups).toHaveLength(1);
    expect(groups[0].template.targetSheet).toBe('January');
    expect(groups[0].clones.map((c) => c.targetSheet)).toEqual(MONTHS.slice(1));
  });

  it('leaves out a subtask whose work differs beyond the sheet name', () => {
    const wave = MONTHS.map((m, i) => monthSubtask(m, i, m === 'June' ? ' Also add a summary row.' : ''));
    const groups = findCloneGroups(wave);
    expect(groups[0].clones.map((c) => c.targetSheet)).not.toContain('June');
    expect(groups[0].clones).toHaveLength(10);
  });

  it('does not group subtasks with different dependencies', () => {
    const wave = MONTHS.slice(0, 4).map((m, i) => ({
      ...monthSubtask(m, i),
      dependsOn: [`dep${i}`],
    }));
    expect(findCloneGroups(wave)).toEqual([]);
  });

  it('does not bother below the minimum group size', () => {
    expect(findCloneGroups(MONTHS.slice(0, 2).map((m, i) => monthSubtask(m, i)))).toEqual([]);
  });

  it('never treats two steps on the SAME sheet as siblings', () => {
    const wave = [0, 1, 2].map((i) => ({ ...monthSubtask('January', i), id: `s${i}` }));
    expect(findCloneGroups(wave)).toEqual([]);
  });
});

describe('cloneActionsForSheet', () => {
  const template = [
    { type: 'ADD_SHEET', sheetName: 'January' },
    { type: 'CREATE_TABLE', sheetName: 'January', range: 'A1:M2', tableName: 'tblJan', hasHeaders: true },
    {
      type: 'BATCH_SET',
      sheetName: 'January',
      operations: [
        { address: 'A1', value: 'Unit No' },
        { address: 'H2', formula: "=IF('January'!F2=\"\",\"\",'January'!F2*G2)" },
        { address: 'J2', value: 'Market rate' },
      ],
    },
    { type: 'SET_DATA_VALIDATION', sheetName: 'January', range: 'I2:I500', source: '=Lists!$B$3:$B$20' },
  ] as unknown as Action[];

  it('re-targets sheetName, table names and sheet-qualified references', () => {
    const [sheet, table, batch, validation] = cloneActionsForSheet(template, 'January', 'February') as unknown as Array<Record<string, any>>;
    expect(sheet.sheetName).toBe('February');
    expect(table.sheetName).toBe('February');
    expect(table.tableName).toBe('tblFeb');
    expect(batch.operations[1].formula).toContain("'February'!F2");
    expect(batch.operations[1].formula).not.toContain('January');
    expect(validation.sheetName).toBe('February');
  });

  it('leaves other sheets untouched', () => {
    const validation = cloneActionsForSheet(template, 'January', 'February')[3] as unknown as Record<string, any>;
    expect(validation.source).toBe('=Lists!$B$3:$B$20');
  });

  it('does not rewrite the abbreviation inside unrelated words', () => {
    const march = [{ type: 'BATCH_SET', sheetName: 'March', operations: [{ address: 'A1', value: 'Market rate' }, { address: 'B1', value: 'tblMar' }] }] as unknown as Action[];
    const out = cloneActionsForSheet(march, 'March', 'April')[0] as unknown as Record<string, any>;
    expect(out.operations[0].value).toBe('Market rate');
    expect(out.operations[1].value).toBe('tblApr');
  });

  it('does not mutate the template actions', () => {
    const snapshot = JSON.stringify(template);
    cloneActionsForSheet(template, 'January', 'February');
    expect(JSON.stringify(template)).toBe(snapshot);
  });
});
