import { ChangeSetController } from '../src/audit/change-set.controller';
import { ChangeSetService } from '../src/audit/change-set.service';
import { WorkbookContext } from '../src/agents/types/agent.types';

/**
 * TASKS.md #250 follow-up — /audit/preview-local lets the client register a
 * change set (and therefore get a working Revert) for actions it already
 * resolved and applied itself with zero LLM calls (the local sheet-action
 * fast lanes). Confirms the controller is a thin, correctly-wired passthrough
 * to the same createPreview() every LLM-routed action already uses.
 */
describe('ChangeSetController.previewLocal (#250 follow-up)', () => {
  const context: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:A1',
        rowCount: 1,
        columnCount: 1,
        values: [['x']],
        formulas: [['']],
        numberFormats: [['General']],
        structure: 'data_table',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('passes the body straight through to createPreview and returns its result', async () => {
    const createPreview = jest.fn().mockResolvedValue({ changeSetId: 'cs_123' });
    const service = { createPreview } as unknown as ChangeSetService;
    const controller = new ChangeSetController(service);

    const result = await controller.previewLocal({
      conversationId: 'conv_1',
      traceId: 'trace_1',
      prompt: 'Copy the Purchase Register sheet and name it March Copy',
      context,
      actions: [{ type: 'ADD_SHEET', name: 'March Copy', copyFrom: 'Sheet1' } as never],
      workbookId: 'wb_1',
    });

    expect(createPreview).toHaveBeenCalledWith({
      conversationId: 'conv_1',
      traceId: 'trace_1',
      prompt: 'Copy the Purchase Register sheet and name it March Copy',
      context,
      actions: [{ type: 'ADD_SHEET', name: 'March Copy', copyFrom: 'Sheet1' }],
      workbookId: 'wb_1',
    });
    expect(result).toEqual({ changeSet: { changeSetId: 'cs_123' } });
  });

  it('defaults traceId to "-" and workbookId to undefined when omitted', async () => {
    const createPreview = jest.fn().mockResolvedValue({ changeSetId: 'cs_456' });
    const service = { createPreview } as unknown as ChangeSetService;
    const controller = new ChangeSetController(service);

    await controller.previewLocal({
      conversationId: 'conv_2',
      prompt: 'Rename the Sheet1 sheet to Renamed',
      context,
      actions: [{ type: 'RENAME_SHEET', oldName: 'Sheet1', newName: 'Renamed' } as never],
    });

    expect(createPreview).toHaveBeenCalledWith(
      expect.objectContaining({ traceId: '-', workbookId: undefined }),
    );
  });
});
