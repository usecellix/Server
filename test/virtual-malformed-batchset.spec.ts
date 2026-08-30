import { virtualApply } from '../src/virtual/virtualApply';
import { buildShadowWorkbook } from '../src/virtual/shadowWorkbook';
import { Action } from '../src/agents/types/agent.types';

/**
 * TASKS.md #156 — a malformed action must not destroy the whole run.
 *
 * Caught by a live smoke test: the Executor emitted a `BATCH_SET` whose
 * operations had no `address`, `virtualBatchSet` passed it to `virtualSetCell`,
 * and `addr.replace(...)` threw "Cannot read properties of undefined (reading
 * 'replace')" — killing the entire agentic run mid-loop, before any Accept card
 * existed. The user saw a bare error and an empty workbook.
 *
 * That the model guesses BATCH_SET's shape is TASKS.md #131 (it is advertised
 * by name with no schema). That a wrong guess is FATAL is this test's concern:
 * the shadow workbook exists to catch bad output, so it must never be the thing
 * that explodes on it.
 */
function baseContext() {
  return {
    activeSheetName: 'Main',
    sheets: [{ name: 'Main', values: [['']], formulas: [['']] }],
    namedRanges: [],
    tables: [],
  } as never;
}

describe('virtualApply — malformed BATCH_SET (TASKS.md #156)', () => {
  it('survives operations with no address at all', () => {
    const wb = buildShadowWorkbook(baseContext());
    const action = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [{ value: 'Month' }, { value: 'Unit No' }],
    } as unknown as Action;

    expect(() => virtualApply(wb, [action])).not.toThrow();
  });

  it('still applies the well-formed operations alongside malformed ones', () => {
    const wb = buildShadowWorkbook(baseContext());
    const action = {
      type: 'BATCH_SET',
      sheetName: 'Main',
      operations: [
        { address: 'A18', value: 'Month' },
        { value: 'dropped — no address' },
        { address: '', value: 'dropped — blank address' },
        { address: 'B18', value: 'Unit No' },
      ],
    } as unknown as Action;

    const out = virtualApply(wb, [action]);
    const main = out.sheets.get('Main');
    expect(main?.cells.get('A18')?.value).toBe('Month');
    expect(main?.cells.get('B18')?.value).toBe('Unit No');
  });

  it('does not poison the shadow for actions that follow it', () => {
    const wb = buildShadowWorkbook(baseContext());
    const actions = [
      { type: 'BATCH_SET', sheetName: 'Main', operations: [{ value: 'no address' }] },
      { type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'still works' },
    ] as unknown as Action[];

    const out = virtualApply(wb, actions);
    expect(out.sheets.get('Main')?.cells.get('A1')?.value).toBe('still works');
  });
});
