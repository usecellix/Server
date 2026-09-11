import {
  isEarlyEmittable,
  selectEarlyEmittable,
  splitEarlyByPhase,
  excludeAlreadyEmitted,
  keysFor,
} from '../src/excel-ai/utils/progressive-emit.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';

/**
 * TASKS.md #174 — progressive emission must never promise something a later
 * pass will rewrite.
 */
const a = (over: Partial<SheetActionPayload>): SheetActionPayload =>
  ({ sheetName: 'January', ...over }) as SheetActionPayload;

describe('isEarlyEmittable', () => {
  it('allows sheet creation and plain content', () => {
    for (const type of ['ADD_SHEET', 'CREATE_SHEET', 'SET_CELL', 'BATCH_SET', 'ADD_ROW']) {
      expect(isEarlyEmittable(a({ type: type as never }))).toBe(true);
    }
  });

  it('withholds SET_FORMULA — the consolidation pass rewrites exactly those', () => {
    expect(isEarlyEmittable(a({ type: 'SET_FORMULA' }))).toBe(false);
  });

  it('withholds CREATE_CHART — the chart pass moves its anchor', () => {
    expect(isEarlyEmittable(a({ type: 'CREATE_CHART' }))).toBe(false);
  });

  it('withholds formatting and layout — the presentation pass invents those', () => {
    for (const type of ['FORMAT_RANGE', 'FREEZE_PANES', 'AUTOFIT_COLUMNS', 'SET_ROW_HEIGHT']) {
      expect(isEarlyEmittable(a({ type: type as never }))).toBe(false);
    }
  });

  it('withholds anything it does not recognise, rather than guessing', () => {
    expect(isEarlyEmittable(a({ type: 'SOME_FUTURE_TYPE' as never }))).toBe(false);
  });
});

describe('selectEarlyEmittable', () => {
  it('keeps wave order and drops the rest', () => {
    const wave = [
      a({ type: 'ADD_SHEET', sheetName: 'January' }),
      a({ type: 'FORMAT_RANGE', row: 0, col: 0 }),
      a({ type: 'SET_CELL', row: 1, col: 0, value: 'x' }),
    ];
    const early = selectEarlyEmittable(wave);
    expect(early.map((x) => x.type)).toEqual(['ADD_SHEET', 'SET_CELL']);
  });

  it('handles an empty or missing wave', () => {
    expect(selectEarlyEmittable([])).toEqual([]);
    expect(selectEarlyEmittable(undefined as never)).toEqual([]);
  });
});

describe('excludeAlreadyEmitted', () => {
  it('removes an action already shown, matching structurally not referentially', () => {
    // sanitize rewrites objects, so the final action is a different instance.
    const shown = a({ type: 'ADD_SHEET', sheetName: 'January' });
    const rewritten = a({ type: 'ADD_SHEET', sheetName: 'January' });
    const out = excludeAlreadyEmitted([rewritten], keysFor([shown]));
    expect(out).toEqual([]);
  });

  it('keeps everything when nothing was emitted early', () => {
    const list = [a({ type: 'SET_CELL', row: 0, col: 0 })];
    expect(excludeAlreadyEmitted(list, [])).toBe(list);
  });

  it('consumes each key ONCE — a genuine duplicate write still survives', () => {
    // Over-filtering would silently drop a real write; that is the expensive
    // direction, so the counting is deliberate.
    const one = a({ type: 'SET_CELL', row: 0, col: 0, value: 'x' });
    const two = a({ type: 'SET_CELL', row: 0, col: 0, value: 'x' });
    const out = excludeAlreadyEmitted([one, two], keysFor([one]));
    expect(out).toHaveLength(1);
  });

  it('never drops a pass-invented action that was never shown', () => {
    const shown = a({ type: 'ADD_SHEET', sheetName: 'January' });
    const invented = a({ type: 'FORMAT_RANGE', row: 0, col: 0 });
    const out = excludeAlreadyEmitted([shown, invented], keysFor([shown]));
    expect(out.map((x) => x.type)).toEqual(['FORMAT_RANGE']);
  });

  it('distinguishes the same action type on different sheets', () => {
    const jan = a({ type: 'ADD_SHEET', sheetName: 'January' });
    const feb = a({ type: 'ADD_SHEET', sheetName: 'February' });
    const out = excludeAlreadyEmitted([jan, feb], keysFor([jan]));
    expect(out.map((x) => x.sheetName)).toEqual(['February']);
  });
});

describe('splitEarlyByPhase — TASKS.md #175', () => {
  it('never emits creates and content in one card', () => {
    // The live failure: a card labelled "Create and fill 13 sheets" carried 12
    // ADD_SHEETs plus every header write, so one occupied A1 blocked all 69
    // actions with no way forward.
    const groups = splitEarlyByPhase([
      a({ type: 'ADD_SHEET', sheetName: 'January' }),
      a({ type: 'BATCH_SET', sheetName: 'January' }),
      a({ type: 'ADD_SHEET', sheetName: 'February' }),
    ]);
    expect(groups).toHaveLength(2);
    expect(groups[0].every((x) => x.type === 'ADD_SHEET')).toBe(true);
    expect(groups[1].every((x) => x.type === 'BATCH_SET')).toBe(true);
  });

  it('puts creates FIRST so content never precedes the sheet it writes into', () => {
    const groups = splitEarlyByPhase([
      a({ type: 'BATCH_SET', sheetName: 'January' }),
      a({ type: 'ADD_SHEET', sheetName: 'January' }),
    ]);
    expect(groups[0][0].type).toBe('ADD_SHEET');
  });

  it('returns a single group when a wave is all one phase', () => {
    expect(splitEarlyByPhase([a({ type: 'ADD_SHEET' })])).toHaveLength(1);
    expect(splitEarlyByPhase([a({ type: 'SET_CELL' })])).toHaveLength(1);
  });

  it('returns nothing for an empty wave', () => {
    expect(splitEarlyByPhase([])).toEqual([]);
  });
});
