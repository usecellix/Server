import {
  buildPendingWritePlanMetadata,
  buildResumedWritePrompt,
  findPendingWritePlan,
  isAffirmationMessage,
  isConfirmationOfferText,
  localActionWithoutLlmMessage,
  localWriteUnavailableMessage,
  shouldStorePendingWritePlan,
} from '../src/excel-ai/utils/pending-write-plan.util';
import type { ConversationMessageEntry } from '../src/excel-ai/schemas/conversation.schema';
import { ConversationEngineService } from '../src/excel-ai/services/conversation-engine.service';
import { IntentClassifierService } from '../src/excel-ai/services/intent-classifier.service';
import { DataQueryService } from '../src/excel-ai/services/data-query.service';
import { SheetAnalyzerService } from '../src/excel-ai/services/sheet-analyzer.service';

describe('pending-write-plan.util', () => {
  const multiMonthPrompt =
    'Create 12 monthly payment sheets Jan-Dec with Unit No, Guest, check-in/out, Rate, total, source, payment status, bank account, plus a Main dashboard sheet.';

  it('detects short affirmations including "yes do it"', () => {
    expect(isAffirmationMessage('yes')).toBe(true);
    expect(isAffirmationMessage('yes do it')).toBe(true);
    expect(isAffirmationMessage('go ahead')).toBe(true);
    expect(isAffirmationMessage('apply it')).toBe(true);
    expect(isAffirmationMessage(multiMonthPrompt)).toBe(false);
  });

  it('detects confirmation-offer assistant prose', () => {
    expect(
      isConfirmationOfferText(
        'Want me to apply this layout? I would need your confirmation to add columns.',
      ),
    ).toBe(true);
    expect(isConfirmationOfferText('I see 51 rows and 12 columns.')).toBe(false);
  });

  it('finds explicit metadata plans and falls back from confirm text + prior user', () => {
    const withMeta: ConversationMessageEntry[] = [
      {
        id: 'u1',
        role: 'user',
        content: multiMonthPrompt,
        type: 'command',
        timestamp: new Date(),
      },
      {
        id: 'a1',
        role: 'assistant',
        content: 'Want me to apply this multi-sheet build?',
        type: 'answer',
        timestamp: new Date(),
        metadata: buildPendingWritePlanMetadata(multiMonthPrompt, 'Want me to apply?'),
      },
    ];
    expect(findPendingWritePlan(withMeta)?.originalPrompt).toBe(multiMonthPrompt);

    const withoutMeta: ConversationMessageEntry[] = [
      {
        id: 'u1',
        role: 'user',
        content: multiMonthPrompt,
        type: 'command',
        timestamp: new Date(),
      },
      {
        id: 'a1',
        role: 'assistant',
        content:
          'Your Purchase Register does not match hospitality columns. Want me to apply a parallel monthly sheet scaffold? I need your confirmation to add them.',
        type: 'answer',
        timestamp: new Date(),
      },
    ];
    expect(findPendingWritePlan(withoutMeta)?.originalPrompt).toBe(multiMonthPrompt);
  });

  it('builds a resume prompt that forbids re-confirming and prefers new sheets', () => {
    const plan = { originalPrompt: multiMonthPrompt };
    const resumed = buildResumedWritePrompt(plan);
    expect(resumed).toContain('User confirmed');
    expect(resumed).toContain('Do not ask for confirmation again');
    expect(resumed).toContain(multiMonthPrompt);
    expect(resumed).toContain('prefer adding new sheets');
  });

  it('only stores plan for confirm-only answers without actions', () => {
    expect(shouldStorePendingWritePlan('Want me to apply these changes?', false)).toBe(true);
    expect(shouldStorePendingWritePlan('Want me to apply these changes?', true)).toBe(false);
    expect(shouldStorePendingWritePlan('Sorted by Amount.', false)).toBe(false);
  });
});

describe('local decide affirmation + write integrity', () => {
  const multiMonthPrompt =
    'Create Jan-Dec monthly sheets with guest payment columns and a Main dashboard.';

  function makeEngine(): ConversationEngineService {
    const sheetAnalyzer = new SheetAnalyzerService();
    const intentClassifier = new IntentClassifierService();
    const dataQuery = new DataQueryService(sheetAnalyzer);
    // Engine only needs sheetAnalyzer, classifier, dataQuery for decide(); remaining deps unused.
    return new ConversationEngineService(
      sheetAnalyzer,
      { hasLlmProvider: false } as never,
      {} as never,
      {} as never,
      intentClassifier,
      dataQuery,
    );
  }

  const sheetData = [
    ['Date', 'Invoice No', 'Supplier', 'Amount'],
    ['2024-01-01', 'INV-1', 'Acme', 100],
  ];

  const analysis = {
    rowCount: 51,
    columnCount: 12,
    headers: ['Date', 'Invoice No', 'Supplier', 'Amount'],
    headerRowIndex: 0,
    isEmpty: false,
    columnLetters: ['A', 'B', 'C', 'D'],
  };

  it('never returns "I see N rows" for yes do it after a confirm offer', () => {
    const engine = makeEngine();
    const history: ConversationMessageEntry[] = [
      {
        id: 'u1',
        role: 'user',
        content: multiMonthPrompt,
        type: 'command',
        timestamp: new Date(),
      },
      {
        id: 'a1',
        role: 'assistant',
        content:
          'Want me to apply this scaffold? I would need your confirmation to create the monthly sheets.',
        type: 'answer',
        timestamp: new Date(),
        metadata: buildPendingWritePlanMetadata(
          multiMonthPrompt,
          'Want me to apply this scaffold?',
        ),
      },
    ];

    const result = engine.decide('yes do it', sheetData, analysis, history);
    expect(result.kind).toBe('answer');
    if (result.kind === 'answer') {
      expect(result.answer).not.toMatch(/I see \d+ rows/i);
      expect(result.answer).toMatch(/pending request|OPENROUTER_API_KEY|AI write path/i);
      expect(result.answer).toContain('monthly');
    }
  });

  it('never returns sheet census for ACTION when LLM is off', () => {
    const engine = makeEngine();
    const result = engine.decide(
      'create 12 monthly sheets and a main dashboard',
      sheetData,
      analysis,
      [],
    );
    expect(result.kind).toBe('answer');
    if (result.kind === 'answer') {
      expect(result.answer).not.toMatch(/I see \d+ rows/i);
      expect(result.answer).toBe(localActionWithoutLlmMessage());
    }
  });

  it('uses honest pending-plan copy when plan exists but LLM is off', () => {
    const plan = { originalPrompt: multiMonthPrompt };
    expect(localWriteUnavailableMessage(plan)).toContain('pending request');
    expect(localWriteUnavailableMessage(plan)).not.toMatch(/I see \d+ rows/i);
  });
});
