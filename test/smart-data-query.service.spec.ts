import { Test } from '@nestjs/testing';
import { SmartDataQueryService } from '../src/excel-ai/services/smart-data-query.service';
import { OpenRouterService } from '../src/excel-ai/services/openrouter.service';

const mockOpenRouter = {
  complete: jest.fn(),
};

const purchaseRegisterRows = [
  ['Date', 'Voucher No', 'CGST', 'SGST'],
  ['01-04-2024', 'INV-001', '1868.41 Dr', '1868.41 Dr'],
  ['02-04-2024', 'INV-002', '945.00 Dr', '945.00 Dr'],
];

describe('SmartDataQueryService', () => {
  let service: SmartDataQueryService;
  const events: Array<{ event: string; data: Record<string, unknown> }> = [];
  const emit = (event: string, data: Record<string, unknown>) => {
    events.push({ event, data });
  };

  beforeEach(async () => {
    events.length = 0;
    const module = await Test.createTestingModule({
      providers: [
        SmartDataQueryService,
        { provide: OpenRouterService, useValue: mockOpenRouter },
      ],
    }).compile();

    service = module.get(SmartDataQueryService);
    jest.clearAllMocks();
  });

  it('emits thinking and returns an answer', async () => {
    mockOpenRouter.complete.mockResolvedValue('Total CGST is ₹2,813.41');

    const answer = await service.handleQuery(
      'What is the total CGST?',
      purchaseRegisterRows,
      undefined,
      'Purchase register',
      emit,
    );

    expect(answer).toBe('Total CGST is ₹2,813.41');
    expect(events.some((entry) => entry.event === 'thinking')).toBe(true);
  });

  it('passes sliced data (CGST column only) to OpenRouter', async () => {
    mockOpenRouter.complete.mockResolvedValue('Total CGST is ₹2,813.41');

    await service.handleQuery(
      'What is the total CGST?',
      purchaseRegisterRows,
      undefined,
      'Purchase register',
      emit,
    );

    const callArgs = mockOpenRouter.complete.mock.calls[0][0];
    expect(callArgs.userMessage).toContain('CGST');
    expect(callArgs.userMessage).toContain('1868.41 Dr');
    expect(callArgs.systemPrompt).toContain('do NOT suggest formulas');
    expect(callArgs.tier).toBe('medium');
    expect(callArgs.responseFormat).toBe('text');
    expect(callArgs.reasoningEffort).toBe('low');
  });

  it('returns fallback answer when sheet data is missing', async () => {
    const answer = await service.handleQuery(
      'What is the total CGST?',
      [],
      undefined,
      undefined,
      emit,
    );

    expect(answer).toContain('could not find any sheet data');
    expect(mockOpenRouter.complete).not.toHaveBeenCalled();
  });

  it('returns error answer when LLM throws', async () => {
    mockOpenRouter.complete.mockRejectedValue(new Error('API timeout'));

    const answer = await service.handleQuery(
      'What is the total CGST?',
      purchaseRegisterRows,
      undefined,
      'Purchase register',
      emit,
    );

    expect(answer).toContain('unable to compute');
  });

  /**
   * TASKS.md #256 — live repro: "How many sheets are in this workbook?"
   * asked right after creating a new blank sheet (making it the active
   * one) hard-failed with "I could not find any sheet data" even though
   * the workbook plainly had 4 sheets — this route is scoped to the
   * active sheet's DATA, but the question needs no data at all, only the
   * sheet list already in `workbookContext`.
   */
  describe('workbook sheet-count/list questions (#256)', () => {
    const fourSheetContext = {
      activeSheet: 'Azhar',
      sheets: [
        { sheetName: 'Purchase Register' },
        { sheetName: 'GSTR-2A' },
        { sheetName: 'Summary' },
        { sheetName: 'Azhar' },
      ],
    } as never;

    it('answers "how many sheets" from the sheet list, with NO sheet data and NO LLM call, even on an empty active sheet', async () => {
      const answer = await service.handleQuery(
        'How many sheets are in this workbook?',
        [], // the active sheet (Azhar) is blank — this used to be a hard failure
        fourSheetContext,
        'Azhar',
        emit,
      );

      expect(answer).toBe(
        'This workbook has 4 sheets: "Purchase Register", "GSTR-2A", "Summary", "Azhar". No hidden sheets.',
      );
      expect(mockOpenRouter.complete).not.toHaveBeenCalled();
    });

    /**
     * TASKS.md #257 — live follow-up: `isHidden` was being read correctly
     * client-side but silently dropped before reaching the backend, so a
     * 6-sheet workbook with 2 real hidden sheets got reported as "6 sheets"
     * with no distinction — the user had to point out "there is a hidden
     * sheet, so mention that clearly."
     */
    it('separates visible from hidden sheets when isHidden is known (#257)', async () => {
      const sixSheetContext = {
        activeSheet: 'Apr 2024 Data',
        sheets: [
          { sheetName: 'Purchase Register', isHidden: false },
          { sheetName: 'GSTR-2A', isHidden: false },
          { sheetName: 'Summary', isHidden: false },
          { sheetName: 'Working', isHidden: true },
          { sheetName: 'StateCodes', isHidden: true },
          { sheetName: 'Apr 2024 Data', isHidden: false },
        ],
      } as never;

      const answer = await service.handleQuery(
        'How many sheets are in this workbook?',
        [],
        sixSheetContext,
        'Apr 2024 Data',
        emit,
      );

      expect(answer).toBe(
        'This workbook has 6 sheets total — 4 visible: "Purchase Register", "GSTR-2A", "Summary", "Apr 2024 Data"; ' +
          '2 hidden: "Working", "StateCodes".',
      );
    });

    it('treats an unknown (undefined) isHidden as visible rather than guessing it is hidden (#257)', async () => {
      const legacyContext = {
        activeSheet: 'Sheet1',
        sheets: [{ sheetName: 'Sheet1' }, { sheetName: 'Sheet2' }],
      } as never;

      const answer = await service.handleQuery(
        'How many sheets are in this workbook?',
        [],
        legacyContext,
        'Sheet1',
        emit,
      );

      expect(answer).toContain('No hidden sheets.');
      expect(answer).not.toContain('hidden:');
    });

    it('also answers "what sheets" / "list the sheets" phrasing', async () => {
      const a1 = await service.handleQuery(
        'What sheets does this workbook have?',
        [],
        fourSheetContext,
        'Azhar',
        emit,
      );
      expect(a1).toContain('4 sheets');

      const a2 = await service.handleQuery('List the sheets', [], fourSheetContext, 'Azhar', emit);
      expect(a2).toContain('4 sheets');
    });

    it('falls through to the normal data-query path (and its own fallback) when no sheet list is available', async () => {
      const answer = await service.handleQuery(
        'How many sheets are in this workbook?',
        [],
        undefined,
        undefined,
        emit,
      );
      expect(answer).toContain('could not find any sheet data');
    });

    it('does not intercept an unrelated data question that merely mentions "sheet"', async () => {
      mockOpenRouter.complete.mockResolvedValue('Total CGST is ₹2,813.41');
      const answer = await service.handleQuery(
        'What is the total CGST on this sheet?',
        purchaseRegisterRows,
        fourSheetContext,
        'Purchase Register',
        emit,
      );
      expect(answer).toBe('Total CGST is ₹2,813.41');
      expect(mockOpenRouter.complete).toHaveBeenCalled();
    });
  });
});
