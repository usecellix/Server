import { splitIntoActionWaves } from '../src/excel-ai/utils/action-wave.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

describe('splitIntoActionWaves', () => {
  it('does not split a pure-write batch (the common case) — single wave, actions unchanged', () => {
    const actions: SheetActionPayload[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', row: 0, col: 0, value: 'x' },
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 0, formula: '=A1' },
    ];
    const waves = splitIntoActionWaves(actions);
    expect(waves).toHaveLength(1);
    // Small batches stay one card — staging a two-cell edit is ceremony, not
    // review. Contents preserved, order preserved. TASKS.md #160.
    expect(waves[0].actions).toEqual(actions);
  });

  it('does not split a pure-structural batch — single wave', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'ADD_SHEET', name: 'February' },
    ];
    const waves = splitIntoActionWaves(actions);
    expect(waves).toHaveLength(1);
    expect(waves[0].label).toBe('Create 2 sheets');
  });

  // TASKS.md #141 — one Accept per build. The two-wave split showed a second,
  // disabled Accept card behind "Accept the earlier step first"; a user who
  // accepted only the first was left with 13 empty tabs and no sign the build
  // was half-done. Ordering (creates first) is kept — it is what makes the
  // writes land — but it now happens inside a single wave.
  it('hoists sheet-creates to the front of ONE wave, never a second card', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
      { type: 'ADD_SHEET', name: 'February' },
      { type: 'SET_CELL', sheetName: 'February', row: 0, col: 0, value: 'Unit No' },
      { type: 'SET_FORMULA', sheetName: 'Main', row: 0, col: 1, formula: '=SUM(January!G:G)' },
    ];
    const waves = splitIntoActionWaves(actions);

    expect(waves).toHaveLength(1);
    expect(waves[0].actions).toHaveLength(5);
    // Every create precedes every write — a write to a sheet created later in
    // the array would fail at apply.
    const lastCreate = waves[0].actions.map((a) => a.type).lastIndexOf('ADD_SHEET');
    const firstWrite = waves[0].actions.findIndex((a) => a.type !== 'ADD_SHEET');
    expect(lastCreate).toBeLessThan(firstWrite);
  });

  it('every action from the input appears in exactly one wave, order preserved within each', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'June' },
      { type: 'SET_CELL', sheetName: 'June', row: 0, col: 0, value: 'a' },
      { type: 'RENAME_SHEET', oldName: 'Sheet1', newName: 'Main' },
      { type: 'SET_FORMULA', sheetName: 'Main', row: 0, col: 0, formula: '=1' },
    ];
    const waves = splitIntoActionWaves(actions);
    const flattened = waves.flatMap((w) => w.actions);
    expect(flattened).toEqual([actions[0], actions[2], actions[1], actions[3]]);
  });

  it('treats CREATE_TABLE/DEFINE_NAMED_RANGE as structural too, without miscounting the sheet-create label', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'Data' },
      { type: 'CREATE_TABLE', sheetName: 'Data', range: 'A1:C1' },
      { type: 'SET_CELL', sheetName: 'Data', row: 1, col: 0, value: 'x' },
    ];
    const waves = splitIntoActionWaves(actions);
    // Three actions is under the staging threshold, so still one card — but
    // the phase ORDER must hold: create, then content, then layout.
    expect(waves).toHaveLength(1);
    expect(waves[0].actions.map((a) => a.type)).toEqual([
      'ADD_SHEET',
      'SET_CELL',
      'CREATE_TABLE',
    ]);
    expect(waves[0].label).toBe('3 changes ready for review');
  });

  it('handles an empty action list without throwing', () => {
    expect(splitIntoActionWaves([])).toEqual([
      { actions: [], label: '0 changes ready for review', actionIndexes: [] },
    ]);
  });
});

/**
 * TASKS.md #160 — staged, individually-acceptable steps.
 *
 * This deliberately reverses #141's single-Accept merge. #141's diagnosis was
 * right (accepting one of two cards left a half-built workbook looking
 * finished) but its remedy was too blunt: the problem was that staging was
 * INVISIBLE, not that it existed. Steps are back, with position reported.
 */
describe('splitIntoActionWaves — staged steps for a large build (TASKS.md #160)', () => {
  function ledgerBuild(): SheetActionPayload[] {
    const a: SheetActionPayload[] = [];
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
    for (const m of months) a.push({ type: 'ADD_SHEET', name: m, sheetName: m });
    for (const m of months) {
      for (let c = 0; c < 5; c += 1) {
        a.push({ type: 'SET_CELL', sheetName: m, row: 0, col: c, value: `H${c}` });
      }
    }
    for (let r = 0; r < 6; r += 1) {
      a.push({ type: 'SET_FORMULA', sheetName: 'Main', row: r, col: 1, formula: '=SUM(Jan!A:A)' });
    }
    for (const m of months) {
      a.push({ type: 'FORMAT_RANGE', sheetName: m, row: 0, col: 0, rowCount: 1, colCount: 5, format: { bold: true } });
      a.push({ type: 'FREEZE_PANES', sheetName: m, freezeRows: 1 });
    }
    a.push({ type: 'CREATE_CHART', sheetName: 'Main', sourceRange: 'A1:B6' });
    return a;
  }

  it('splits a large build into several ordered steps', () => {
    const waves = splitIntoActionWaves(ledgerBuild());
    expect(waves.length).toBeGreaterThan(1);
    expect(waves.length).toBeLessThanOrEqual(6);
  });

  it('every action survives in exactly one step', () => {
    const input = ledgerBuild();
    const flattened = splitIntoActionWaves(input).flatMap((w) => w.actions);
    expect(flattened).toHaveLength(input.length);
    expect(new Set(flattened).size).toBe(input.length);
  });

  it('orders steps so each is safe to apply on its own', () => {
    const waves = splitIntoActionWaves(ledgerBuild());
    const phaseOf = (t: string) =>
      ['ADD_SHEET'].includes(t) ? 0
      : ['SET_CELL'].includes(t) ? 1
      : ['SET_FORMULA'].includes(t) ? 2
      : ['FORMAT_RANGE'].includes(t) ? 3
      : ['FREEZE_PANES'].includes(t) ? 4
      : 5;
    // Phase index must never decrease across the flattened sequence: sheets
    // before writes, writes before formulas, content before autofit/freeze,
    // everything before charts.
    const seq = waves.flatMap((w) => w.actions).map((a) => phaseOf(a.type));
    for (let i = 1; i < seq.length; i += 1) expect(seq[i]).toBeGreaterThanOrEqual(seq[i - 1]);
  });

  it('puts every sheet-create in the first step, before any write', () => {
    const waves = splitIntoActionWaves(ledgerBuild());
    expect(waves[0].actions.every((a) => a.type === 'ADD_SHEET')).toBe(true);
    expect(waves[0].label).toBe('Create 6 sheets');
    const laterCreates = waves.slice(1).flatMap((w) => w.actions).filter((a) => a.type === 'ADD_SHEET');
    expect(laterCreates).toHaveLength(0);
  });

  it('gives each step a human label, not an action dump', () => {
    const labels = splitIntoActionWaves(ledgerBuild()).map((w) => w.label);
    expect(labels[0]).toBe('Create 6 sheets');
    expect(labels.join(' ')).toMatch(/Write content|Add formulas|Apply formatting|Finish layout|Add charts/);
  });

  it('does NOT stage a small batch — that would be ceremony, not review', () => {
    const small: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'Notes', sheetName: 'Notes' },
      { type: 'SET_CELL', sheetName: 'Notes', row: 0, col: 0, value: 'Title' },
      { type: 'FORMAT_RANGE', sheetName: 'Notes', row: 0, col: 0, rowCount: 1, colCount: 1, format: { bold: true } },
    ];
    expect(splitIntoActionWaves(small)).toHaveLength(1);
  });

  it('never exceeds the step budget, merging rather than dropping', () => {
    const many: SheetActionPayload[] = [];
    for (let i = 0; i < 10; i += 1) many.push({ type: 'ADD_SHEET', name: `S${i}`, sheetName: `S${i}` });
    for (let i = 0; i < 10; i += 1) many.push({ type: 'SET_CELL', sheetName: 'S0', row: i, col: 0, value: 'x' });
    for (let i = 0; i < 10; i += 1) many.push({ type: 'SET_FORMULA', sheetName: 'S0', row: i, col: 1, formula: '=1' });
    for (let i = 0; i < 10; i += 1) many.push({ type: 'FORMAT_RANGE', sheetName: 'S0', row: i, col: 0, rowCount: 1, colCount: 1, format: { bold: true } });
    for (let i = 0; i < 10; i += 1) many.push({ type: 'FREEZE_PANES', sheetName: `S${i}`, freezeRows: 1 });
    many.push({ type: 'CREATE_CHART', sheetName: 'S0', sourceRange: 'A1:B2' });

    const waves = splitIntoActionWaves(many);
    expect(waves.length).toBeLessThanOrEqual(6);
    expect(waves.flatMap((w) => w.actions)).toHaveLength(many.length);
  });

  it('is domain-agnostic — same staging for a non-calendar build', () => {
    const a: SheetActionPayload[] = [];
    const regions = ['North', 'South', 'East', 'West', 'Central', 'Overseas'];
    for (const r of regions) a.push({ type: 'ADD_SHEET', name: r, sheetName: r });
    for (const r of regions) {
      for (let c = 0; c < 4; c += 1) a.push({ type: 'SET_CELL', sheetName: r, row: 0, col: c, value: `C${c}` });
    }
    for (const r of regions) a.push({ type: 'AUTOFIT_COLUMNS', sheetName: r });
    const waves = splitIntoActionWaves(a);
    expect(waves.length).toBeGreaterThan(1);
    expect(waves[0].label).toBe('Create 6 sheets');
  });
});
