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
});
