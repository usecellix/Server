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
    expect(waves[0].actions).toBe(actions);
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

  it('splits sheet-creates from everything else, creates first', () => {
    const actions: SheetActionPayload[] = [
      { type: 'ADD_SHEET', name: 'January' },
      { type: 'SET_CELL', sheetName: 'January', row: 0, col: 0, value: 'Unit No' },
      { type: 'ADD_SHEET', name: 'February' },
      { type: 'SET_CELL', sheetName: 'February', row: 0, col: 0, value: 'Unit No' },
      { type: 'SET_FORMULA', sheetName: 'Main', row: 0, col: 1, formula: '=SUM(January!G:G)' },
    ];
    const waves = splitIntoActionWaves(actions);

    expect(waves).toHaveLength(2);
    expect(waves[0].actions.every((a) => a.type === 'ADD_SHEET')).toBe(true);
    expect(waves[0].actions).toHaveLength(2);
    expect(waves[1].actions).toHaveLength(3);
    expect(waves[1].actions.some((a) => a.type === 'ADD_SHEET')).toBe(false);
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
    expect(waves).toHaveLength(2);
    expect(waves[0].actions).toHaveLength(2);
    // Mixed structural (1 sheet-create + 1 table) — falls back to the generic label.
    expect(waves[0].label).toBe('2 changes ready for review');
  });

  it('handles an empty action list without throwing', () => {
    expect(splitIntoActionWaves([])).toEqual([{ actions: [], label: '0 changes ready for review' }]);
  });
});
