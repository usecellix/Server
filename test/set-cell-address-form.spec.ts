import { normalizeSingleAction } from '../src/agents/utils/normalize-executor-output.util';
import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';
import { SheetAnalysis } from '../src/excel-ai/services/sheet-analyzer.service';

/**
 * TASKS.md #183 — "add a row with 5000 in the amount column" planned, executed
 * and verified cleanly ("2 actions ready for preview"), then died with
 * "Something went wrong applying this change — try rephrasing".
 *
 * The Executor emitted its two SET_CELLs in **address** form (`"address": "A1"`)
 * rather than row/col — the shape executor.prompt.ts itself demonstrates for
 * BATCH_SET operations immediately above the line telling the model to "emit
 * individual SET_CELL actions instead" when unsure. Nothing upstream minds:
 * `normalizeSingleAction` copies `address` through, `hasRequiredFields` only
 * guards BATCH_SET/CONDITIONAL_FORMAT, and the verifier passes. Then
 * `sanitizeAction` demands integer row/col for SET_CELL, returns null for both,
 * and `finalizeActions` hands back an empty array.
 *
 * `expandRangeStringToIndices` already exists for exactly this reason — but it
 * only reads `range`, and only for the range-shaped action types. The
 * cell-addressed family (SET_CELL / SET_FORMULA / CLEAR_CELL / HIGHLIGHT_CELL)
 * had the same gap one field over.
 */
describe('SET_CELL emitted in address form (TASKS.md #183)', () => {
  const service = new ConversationEngineService(
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
  );

  const emptySheet: SheetAnalysis = {
    rowCount: 0,
    columnCount: 0,
    headers: [],
    isEmpty: true,
    columnLetters: [],
    headerRowIndex: 0,
  };

  it('resolves address to row/col during executor normalization', () => {
    const action = normalizeSingleAction(
      { type: 'SET_CELL', sheetName: 'Main', address: 'A1', value: 'Amount' },
      'Main',
    );

    expect(action).not.toBeNull();
    expect(action).toMatchObject({ type: 'SET_CELL', row: 0, col: 0, value: 'Amount' });
  });

  it('resolves a SET_FORMULA address too', () => {
    const action = normalizeSingleAction(
      { type: 'SET_FORMULA', sheetName: 'Main', address: 'C3', formula: '=SUM(A1:A2)' },
      'Main',
    );

    expect(action).toMatchObject({ type: 'SET_FORMULA', row: 2, col: 2 });
  });

  it('keeps an explicit row/col pair over the address when both are present', () => {
    const action = normalizeSingleAction(
      { type: 'SET_CELL', sheetName: 'Main', address: 'Z99', row: 4, col: 1, value: 7 },
      'Main',
    );

    expect(action).toMatchObject({ row: 4, col: 1 });
  });

  it('survives finalizeActions instead of being silently dropped', () => {
    const actions = [
      { type: 'SET_CELL', sheetName: 'Main', address: 'A1', value: 'Amount' },
      { type: 'SET_CELL', sheetName: 'Main', address: 'A2', value: 5000 },
    ].map((raw) => normalizeSingleAction(raw, 'Main')!);

    expect(actions.every(Boolean)).toBe(true);

    // The live shape: the workbook context names the (empty) Main sheet, so
    // the header guard knows there is no header row to protect and leaves both
    // writes exactly as the Executor meant them.
    const finalized = service.finalizeActions(
      actions,
      emptySheet,
      {
        activeSheet: 'Main',
        sheets: [{ sheetName: 'Main', headers: [], rowCount: 0 }],
      } as never,
      'add a row with 5000 in the amount column',
    );

    expect(finalized).toHaveLength(2);
    expect(finalized.map((a) => a.type)).toEqual(['SET_CELL', 'SET_CELL']);
    expect(finalized[0]).toMatchObject({ row: 0, col: 0, value: 'Amount' });
    expect(finalized[1]).toMatchObject({ row: 1, col: 0, value: 5000 });
  });

  // Without the address resolution these two produced an empty batch, which
  // conversation.service.ts reports as WriteRouteNoActionError — the user-facing
  // "Something went wrong applying this change — try rephrasing".
  it('does not silently reduce an address-form batch to nothing', () => {
    const actions = [
      { type: 'SET_CELL', sheetName: 'Main', address: 'A1', value: 'Amount' },
      { type: 'SET_CELL', sheetName: 'Main', address: 'A2', value: 5000 },
    ].map((raw) => normalizeSingleAction(raw, 'Main')!);

    const finalized = service.finalizeActions(
      actions,
      emptySheet,
      undefined,
      'add a row with 5000 in the amount column',
    );

    expect(finalized.length).toBeGreaterThan(0);
  });
});
