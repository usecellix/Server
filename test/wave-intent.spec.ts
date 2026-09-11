import * as fs from 'fs';
import * as path from 'path';
import {
  attributeActionsToSubtasks,
  intentForWave,
  resolveSheetName,
  PlanIntentEntry,
} from '../src/excel-ai/utils/wave-intent.util';
import { splitIntoActionWaves } from '../src/excel-ai/utils/action-wave.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

const realLedgerActions: SheetActionPayload[] = JSON.parse(
  fs.readFileSync(path.join(__dirname, '__fixtures__real-ledger-actions.json'), 'utf8'),
);

/**
 * TASKS.md #167 — each staged step describes its OWN work.
 *
 * The regression this guards is TASKS.md #161's: every card rendering the same
 * whole-plan bullets, so "Create 13 sheets" promised the chart that arrives
 * five steps later. #161 fixed that by suppressing intent on staged builds;
 * this makes the over-promising impossible by construction instead, which is
 * what #161's own note said should replace it.
 */
describe('attributeActionsToSubtasks', () => {
  it('matches an action the finalize passes left untouched', () => {
    const action = { type: 'ADD_SHEET', sheetName: 'January' } as SheetActionPayload;
    const attribution = attributeActionsToSubtasks(
      [action],
      [{ subtaskId: 's1', actions: [action] }],
    );
    expect(attribution.get(0)).toBe('s1');
  });

  it('matches structurally, not referentially — a rewritten action still attributes', () => {
    // sanitize/consolidation/presentation all REPLACE action objects, so `===`
    // matching would attribute almost nothing.
    const emitted = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [{ address: 'A4', value: 'Month' }],
    } as SheetActionPayload;
    const rewritten = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [
        { address: 'A4', value: 'Month' },
        { address: 'B4', value: 'Total Amount' },
      ],
    } as SheetActionPayload;
    const attribution = attributeActionsToSubtasks(
      [rewritten],
      [{ subtaskId: 's9', actions: [emitted] }],
    );
    expect(attribution.get(0)).toBe('s9');
  });

  it('attributes pass-generated actions by sheet ownership', () => {
    // The presentation pass invents this FORMAT_RANGE; no subtask emitted it,
    // but January's formatting is plainly part of January's work.
    const owned = { type: 'ADD_SHEET', sheetName: 'January' } as SheetActionPayload;
    const invented = {
      type: 'FORMAT_RANGE',
      sheetName: 'January',
      row: 0,
      col: 0,
    } as SheetActionPayload;
    const attribution = attributeActionsToSubtasks(
      [owned, invented],
      [{ subtaskId: 's3', actions: [owned] }],
    );
    expect(attribution.get(1)).toBe('s3');
  });

  it('leaves an action on a sheet nobody owns unattributed', () => {
    const attribution = attributeActionsToSubtasks(
      [{ type: 'FORMAT_RANGE', sheetName: 'Ghost', row: 0, col: 0 } as SheetActionPayload],
      [{ subtaskId: 's1', actions: [{ type: 'ADD_SHEET', sheetName: 'Real' } as SheetActionPayload] }],
    );
    expect(attribution.has(0)).toBe(false);
  });

  it('gives a contested sheet to the subtask that did the most on it', () => {
    const sheet = (type: string) => ({ type, sheetName: 'Main', row: 0, col: 0 }) as SheetActionPayload;
    const attribution = attributeActionsToSubtasks(
      [sheet('FREEZE_PANES')],
      [
        { subtaskId: 'few', actions: [sheet('SET_CELL')] },
        { subtaskId: 'many', actions: [sheet('SET_FORMULA'), sheet('ADD_ROW'), sheet('WRITE_TABLE')] },
      ],
    );
    expect(attribution.get(0)).toBe('many');
  });
});

describe('intentForWave', () => {
  const plan: PlanIntentEntry[] = [
    { id: 's1', description: 'Create the 12 month sheets', targetSheet: 'January' },
    { id: 's2', description: 'Build the Main dashboard', targetSheet: 'Main' },
  ];

  it('returns only the subtasks a wave actually drew from', () => {
    const attribution = new Map([
      [0, 's1'],
      [1, 's2'],
    ]);
    expect(intentForWave([0], attribution, plan)).toEqual([plan[0]]);
    expect(intentForWave([1], attribution, plan)).toEqual([plan[1]]);
  });

  it('keeps plan order when a wave spans several subtasks', () => {
    const attribution = new Map([
      [0, 's2'],
      [1, 's1'],
    ]);
    expect(intentForWave([0, 1], attribution, plan)).toEqual(plan);
  });

  it('returns [] for a wave attributed to nothing, so the caller falls back', () => {
    // NOT an empty card — #149 hit that bug once already.
    expect(intentForWave([7, 8], new Map(), plan)).toEqual([]);
  });
});

describe('#167 — replayed on the real 140-action build', () => {
  // One subtask per sheet, which is the shape the live plan actually had.
  const sheets = [...new Set(realLedgerActions.map(resolveSheetName).filter(Boolean))];
  const completedSubtasks = sheets.map((sheet, i) => ({
    subtaskId: `s${i}`,
    actions: realLedgerActions.filter((a) => resolveSheetName(a) === sheet),
  }));
  const plan: PlanIntentEntry[] = sheets.map((sheet, i) => ({
    id: `s${i}`,
    description:
      sheet === 'Main'
        ? 'Build the Main dashboard with KPIs, monthly totals and the consolidated list'
        : `Create sheet '${sheet}' and set its column headers`,
    targetSheet: sheet,
  }));

  const waves = splitIntoActionWaves(realLedgerActions);
  const attribution = attributeActionsToSubtasks(realLedgerActions, completedSubtasks);

  it('still splits into the six known steps', () => {
    expect(waves).toHaveLength(6);
  });

  it('attributes every action in the real build', () => {
    expect(attribution.size).toBe(realLedgerActions.length);
  });

  it('gives each step a DIFFERENT intent set — the #161 regression', () => {
    const perWave = waves.map((w) =>
      intentForWave(w.actionIndexes, attribution, plan).map((p) => p.id).join(','),
    );
    // The bug was all six cards showing identical bullets.
    expect(new Set(perWave).size).toBeGreaterThan(1);
  });

  it('never lets a step claim a sheet it does not touch', () => {
    for (const wave of waves) {
      const touched = new Set(wave.actions.map(resolveSheetName).filter(Boolean));
      for (const entry of intentForWave(wave.actionIndexes, attribution, plan)) {
        expect(touched.has(entry.targetSheet)).toBe(true);
      }
    }
  });

  it('keeps actionIndexes consistent with actions through every merge', () => {
    const seen = new Set<number>();
    for (const wave of waves) {
      expect(wave.actionIndexes).toHaveLength(wave.actions.length);
      wave.actionIndexes.forEach((index, i) => {
        expect(realLedgerActions[index]).toBe(wave.actions[i]);
        expect(seen.has(index)).toBe(false);
        seen.add(index);
      });
    }
    expect(seen.size).toBe(realLedgerActions.length);
  });
});
