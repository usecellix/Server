import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';

/**
 * TASKS.md #241 — diagnostic only (see the code comment for why this isn't a
 * blocking guard like #238's row-delete one): a live run produced a bare
 * ADD_SHEET for a copy-intent message despite executor.prompt.ts's explicit
 * warning, while two other runs of the identical prompt that same day got it
 * right. This pins that the miss is at least logged, not silent.
 */
describe('copy-intent diagnostic (#241)', () => {
  const engine = new ConversationEngineService(
    undefined as never,
    { hasLlmProvider: false } as never,
    undefined as never,
    undefined as never,
    undefined as never,
    undefined as never,
  );
  const analysis = { isEmpty: false, headers: ['A'], rowCount: 5, columnCount: 1 } as never;

  it('warns when a copy-intent message produced a bare ADD_SHEET', () => {
    const warnSpy = jest.spyOn((engine as unknown as { logger: { warn: jest.Mock } }).logger, 'warn');

    engine.finalizeActions(
      [{ type: 'ADD_SHEET', name: 'March Copy' } as never],
      analysis,
      undefined,
      'Copy the Purchase Register sheet and name it March Copy',
    );

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining('TASKS.md #241'));
  });

  it('does not warn when COPY_SHEET was correctly produced', () => {
    const warnSpy = jest.spyOn((engine as unknown as { logger: { warn: jest.Mock } }).logger, 'warn');
    warnSpy.mockClear();

    engine.finalizeActions(
      [{ type: 'COPY_SHEET', sheetName: 'Purchase Register', newSheetName: 'March Copy' } as never],
      analysis,
      undefined,
      'Copy the Purchase Register sheet and name it March Copy',
    );

    expect(warnSpy).not.toHaveBeenCalledWith(expect.stringContaining('TASKS.md #241'));
  });

  it('does not warn for a plain create-sheet request (no copy intent)', () => {
    const warnSpy = jest.spyOn((engine as unknown as { logger: { warn: jest.Mock } }).logger, 'warn');
    warnSpy.mockClear();

    engine.finalizeActions(
      [{ type: 'ADD_SHEET', name: 'Q1 Summary' } as never],
      analysis,
      undefined,
      'Create a new sheet called Q1 Summary',
    );

    expect(warnSpy).not.toHaveBeenCalledWith(expect.stringContaining('TASKS.md #241'));
  });

  it('does not warn when ADD_SHEET legitimately carries copyFrom', () => {
    const warnSpy = jest.spyOn((engine as unknown as { logger: { warn: jest.Mock } }).logger, 'warn');
    warnSpy.mockClear();

    engine.finalizeActions(
      [{ type: 'ADD_SHEET', name: 'March Copy', copyFrom: 'Purchase Register' } as never],
      analysis,
      undefined,
      'Copy the Purchase Register sheet and name it March Copy',
    );

    expect(warnSpy).not.toHaveBeenCalledWith(expect.stringContaining('TASKS.md #241'));
  });
});
