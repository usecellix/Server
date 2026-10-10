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

    // These amounts carry a "Dr" suffix, so the column is text and cannot be
    // added up in code. The plan is asked for first and discarded; the model
    // then reads the rows, as before. That second call is the one checked here.
    expect(mockOpenRouter.complete).toHaveBeenCalledTimes(2);
    const callArgs = mockOpenRouter.complete.mock.calls[1][0];
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
   * TASKS.md #380 to #383 — a question about the rows is planned by the model
   * and computed in code, so the figures cannot be mis-added or mis-copied.
   */
  describe('questions answered in code', () => {
    const D = (iso: string) => Math.round(Date.parse(`${iso}T00:00:00Z`) / 86_400_000) + 25_569;
    const statement = [
      ['Date', 'Description', 'Debit', 'Credit', 'Balance'],
      [D('2026-04-01'), 'SALARY', '', 50000, 75000],
      [D('2026-04-03'), 'ATM WDL 593111', 6173.55, '', 68826.45],
      [D('2026-04-04'), 'UPI GROCERY', 100, '', 68726.45],
      [D('2026-04-05'), 'UPI CAFE', 200, '', 68526.45],
      [D('2026-05-18'), 'CHQ PAID FUEL STATION', 20000, '', 48526.45],
    ];
    const context = {
      activeSheet: 'Bank Statement',
      sheets: [
        {
          sheetName: 'Bank Statement',
          usedRange: 'A1:E6',
          rowCount: 6,
          colCount: 5,
          headers: statement[0],
          headerRowIndex: 0,
          sampleData: [],
          columnMeta: [
            { index: 0, sampleValues: [], detectedType: 'date', numberFormat: 'dd-mmm-yyyy' },
            { index: 2, sampleValues: [], detectedType: 'number', numberFormat: '#,##0.00' },
          ],
        },
      ],
    } as never;

    it('computes the answer from a plan, with one model call and no rows sent', async () => {
      mockOpenRouter.complete.mockResolvedValue('{"queries":[{"op":"top","column":"Debit","limit":1}]}');

      const pointers: Parameters<SmartDataQueryService['handleQuery']>[5] = {};
      const answer = await service.handleQuery('Which is the largest debit?', statement, context, 'Bank Statement', emit, pointers);

      // The row it names comes back as a pointer that selects the whole row, A6:E6.
      expect(pointers.matches).toEqual([
        { label: 'row 6', sheetName: 'Bank Statement', row: 5, col: 0, colLetter: 'A', rowNum: 6, rawValue: '', endCol: 4, detail: 'A6:E6' },
      ]);
      expect(answer).toBe(
        '**1 largest by Debit**\n1. 18-May-2026 · CHQ PAID FUEL STATION · Debit 20,000.00 (row 6)\n\nWorked out from all 5 rows of Bank Statement.',
      );
      expect(mockOpenRouter.complete).toHaveBeenCalledTimes(1);
      const sent = mockOpenRouter.complete.mock.calls[0][0];
      expect(sent.responseFormat).toBe('json_object');
      // The planner sees column names and three sample cells each, not the rows.
      expect(sent.userMessage).not.toContain('20000');
      expect(sent.userMessage).not.toContain('FUEL STATION');
    });

    it('says so when the sheet reported more rows than could be read', async () => {
      mockOpenRouter.complete.mockResolvedValue('{"queries":[{"op":"count"}]}');
      const bigger = { activeSheet: 'Bank Statement', sheets: [{ ...(context as { sheets: object[] }).sheets[0], rowCount: 835 }] } as never;

      const answer = await service.handleQuery('How many transactions are there?', statement, bigger, 'Bank Statement', emit);

      expect(answer).toContain('**Number of rows: 5**');
      expect(answer).toMatch(/Only 5 rows of 834 on Bank Statement could be read, so this may be incomplete\.$/);
    });

    it('shows dates as dates when it has to fall back to the model reading the rows', async () => {
      mockOpenRouter.complete
        .mockResolvedValueOnce('{"unsupported":"needs judgement"}')
        .mockResolvedValueOnce('It looks seasonal.');

      await service.handleQuery('Which month looks unusual, and why?', statement, context, 'Bank Statement', emit);

      const fallback = mockOpenRouter.complete.mock.calls[1][0];
      expect(fallback.userMessage).toContain('01-Apr-2026');
      expect(fallback.userMessage).not.toContain(String(D('2026-04-01')));
    });

    it('never returns a blank answer when the model returns no text', async () => {
      mockOpenRouter.complete.mockResolvedValueOnce('{"unsupported":"x"}').mockResolvedValueOnce('   ');

      const answer = await service.handleQuery('What is the total of Debit?', statement, context, 'Bank Statement', emit);

      expect(answer).toMatch(/^I could not work out an answer to that from the sheet\./);
    });

    it('gives no pointers when the data does not start at A1, since its rows would point at the wrong place', async () => {
      mockOpenRouter.complete.mockResolvedValue('{"queries":[{"op":"top","column":"Debit","limit":1}]}');
      const shifted = { activeSheet: 'Bank Statement', sheets: [{ ...(context as { sheets: object[] }).sheets[0], usedRange: "'Bank Statement'!B3:F8" }] } as never;
      const pointers: Parameters<SmartDataQueryService['handleQuery']>[5] = {};

      const answer = await service.handleQuery('Which is the largest debit?', statement, shifted, 'Bank Statement', emit, pointers);

      expect(answer).toContain('Debit 20,000.00');
      expect(pointers.matches).toEqual([]);
    });

    it('leaves find and lookup requests on their existing path', async () => {
      mockOpenRouter.complete.mockResolvedValue('Found it on row 3.');

      await service.handleQuery('Find 593111', statement, context, 'Bank Statement', emit);

      expect(mockOpenRouter.complete).toHaveBeenCalledTimes(1);
      expect(mockOpenRouter.complete.mock.calls[0][0].responseFormat).toBe('text');
    });
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
