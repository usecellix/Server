import { WebChatService } from '../src/web-chat/web-chat.service';
import { InsufficientCreditError } from '../src/credit/errors/insufficient-credit.error';

interface StubConversation {
  conversationId: string;
  userId?: string;
  title?: string;
  messages?: Array<{
    id: string;
    role: 'user' | 'assistant';
    content: string;
    metadata?: { actions?: unknown[] };
  }>;
  sheetSnapshot?: { rowCount: number; columnCount: number; headers: string[] };
  updatedAt?: Date;
}

/**
 * Minimal stand-in for the Mongoose model, supporting the two query shapes
 * WebChatService actually issues: `findOne(...).lean()` and
 * `find(...).sort().limit().lean()`.
 */
function conversationModelStub(docs: StubConversation[]) {
  const findOne = jest.fn((filter: Record<string, unknown>) => ({
    lean: () =>
      Promise.resolve(
        docs.find(
          (doc) =>
            doc.conversationId === filter.conversationId && doc.userId === filter.userId,
        ) ?? null,
      ),
  }));

  const find = jest.fn((filter: Record<string, unknown>) => ({
    sort: () => ({
      limit: () => ({
        lean: () => Promise.resolve(docs.filter((doc) => doc.userId === filter.userId)),
      }),
    }),
  }));

  return { findOne, find } as never;
}

function makeService(overrides: {
  docs?: StubConversation[];
  canStart?: boolean;
  debited?: number;
  requested?: number;
  settleResult?: unknown;
  complete?: jest.Mock;
  settle?: jest.Mock;
} = {}) {
  const complete = overrides.complete ?? jest.fn().mockResolvedValue('  An answer.  ');
  const settle =
    overrides.settle ??
    jest.fn().mockResolvedValue(
      overrides.settleResult !== undefined
        ? overrides.settleResult
        : {
            debited: overrides.debited ?? 3,
            requested: overrides.requested ?? overrides.debited ?? 3,
            costUsd: 0.005,
            balances: { planCredits: 100, purchasedCredits: 20, oneTimeCredits: 0 },
          },
    );
  const canStart = jest.fn().mockResolvedValue(overrides.canStart ?? true);

  const service = new WebChatService(
    conversationModelStub(overrides.docs ?? []),
    { complete } as never,
    { canStart, settle } as never,
  );

  return { service, complete, settle, canStart };
}

describe('WebChatService', () => {
  describe('pricing', () => {
    it('prices a short cross-session question as a simple Q&A (2 credits)', () => {
      const { service } = makeService();
      expect(service.estimateCost('what did I do', false)).toEqual({
        complexity: 'simple',
        credits: 2,
      });
    });

    it('prices a synthesis question across sessions as complex (5 credits)', () => {
      const { service } = makeService();
      expect(service.estimateCost('compare my last two reconciliations', false)).toEqual({
        complexity: 'complex',
        credits: 5,
      });
    });

    it('prices a question scoped to ONE session as simple even when its wording looks complex', () => {
      // A single transcript is inherently narrower than a cross-thread
      // synthesis, so scope wins over wording — see classify()'s docblock.
      const { service } = makeService();
      expect(service.estimateCost('compare every row across all my sheets', true)).toEqual({
        complexity: 'simple',
        credits: 2,
      });
    });

    it('estimateCost is a heads-up figure only — ask() bills real usage, not this number', async () => {
      const { service, settle } = makeService({
        docs: [{ conversationId: 'c1', userId: 'u1', title: 'GST' }],
        debited: 4, // a real usage charge that differs from the 5-credit catalog estimate
      });

      const estimate = service.estimateCost('summarise all my sessions', false);
      const answer = await service.ask('u1', 'summarise all my sessions');

      expect(estimate.credits).toBe(5);
      expect(answer.creditsDeducted).toBe(4);
      expect(settle).toHaveBeenCalledTimes(1);
    });
  });

  describe('credit gating', () => {
    it('throws before making any LLM call when the balance is already at 0', async () => {
      const { service, complete, settle } = makeService({ canStart: false });

      await expect(service.ask('u1', 'compare all my sessions')).rejects.toBeInstanceOf(
        InsufficientCreditError,
      );

      // The whole point of gating first (CD-4): no model spend on a request
      // that was never going to be allowed.
      expect(complete).not.toHaveBeenCalled();
      expect(settle).not.toHaveBeenCalled();
    });

    it('never reaches settle when the LLM call fails — no charge for a run that produced nothing', async () => {
      const complete = jest.fn().mockRejectedValue(new Error('provider down'));
      const { service, settle } = makeService({ complete });

      await expect(service.ask('u1', 'what did I do')).rejects.toThrow('provider down');
      expect(settle).not.toHaveBeenCalled();
    });

    it('still returns the answer when the balance runs out mid-settle, rather than discarding paid-for work', async () => {
      // The answer already cost real model spend; a short/zero settle should
      // not throw away work that already happened.
      const { service } = makeService({
        settleResult: { debited: 0, requested: 4, costUsd: 0.006, balances: { planCredits: 0, purchasedCredits: 0, oneTimeCredits: 0 } },
      });

      const answer = await service.ask('u1', 'what did I do');

      expect(answer.answer).toBe('An answer.');
      expect(answer.creditsDeducted).toBe(0);
      expect(answer.newBalance).toBe(0);
    });

    it('reports the post-debit balance so the client can update without a refetch', async () => {
      const { service } = makeService();
      const answer = await service.ask('u1', 'what did I do');
      expect(answer.newBalance).toBe(120);
    });

    it('reports newBalance: 0 rather than throwing when settle itself returns null (e.g. anonymous/zero-cost)', async () => {
      const { service } = makeService({ settleResult: null });
      const answer = await service.ask('u1', 'what did I do');
      expect(answer).toMatchObject({ creditsDeducted: 0, newBalance: 0 });
    });
  });

  describe('context loading', () => {
    it('scopes to one conversation and filters by owner, so another user\'s id reads nothing', async () => {
      const { service, complete } = makeService({
        docs: [{ conversationId: 'c1', userId: 'someone-else', title: 'Their GST work' }],
      });

      const answer = await service.ask('u1', 'what happened', { conversationId: 'c1' });

      expect(answer.citations).toEqual([]);
      // The prompt must not contain the other user's session.
      expect(complete.mock.calls[0][0].userMessage).not.toContain('Their GST work');
    });

    it('cites the sessions it drew on', async () => {
      const { service } = makeService({
        docs: [
          { conversationId: 'c1', userId: 'u1', title: 'GST reconciliation' },
          { conversationId: 'c2', userId: 'u1', title: 'Tally cleanup' },
        ],
      });

      const answer = await service.ask('u1', 'what did I do');

      expect(answer.citations).toEqual([
        { conversationId: 'c1', title: 'GST reconciliation' },
        { conversationId: 'c2', title: 'Tally cleanup' },
      ]);
    });

    it('falls back to a placeholder title rather than citing an empty string', async () => {
      const { service } = makeService({
        docs: [{ conversationId: 'c1', userId: 'u1', title: '   ' }],
      });

      const answer = await service.ask('u1', 'what did I do');
      expect(answer.citations[0].title).toBe('Untitled session');
    });

    it('puts the sheet snapshot in the prompt, so structural questions are answerable without a live workbook', async () => {
      const { service, complete } = makeService({
        docs: [
          {
            conversationId: 'c1',
            userId: 'u1',
            title: 'Ledger',
            sheetSnapshot: { rowCount: 400, columnCount: 6, headers: ['Date', 'GSTIN', 'Amount'] },
          },
        ],
      });

      await service.ask('u1', 'what columns were there');

      const prompt = complete.mock.calls[0][0].userMessage;
      expect(prompt).toContain('400 rows x 6 columns');
      expect(prompt).toContain('Date, GSTIN, Amount');
    });

    it('notes that changes were applied, since "what changed?" lives in metadata not prose', async () => {
      const { service, complete } = makeService({
        docs: [
          {
            conversationId: 'c1',
            userId: 'u1',
            title: 'Build',
            messages: [
              {
                id: 'm1',
                role: 'assistant',
                content: 'Done.',
                metadata: { actions: [{}, {}, {}] },
              },
            ],
          },
        ],
      });

      await service.ask('u1', 'what changed');

      expect(complete.mock.calls[0][0].userMessage).toContain('applied 3 changes');
    });

    it('tells the model plainly when there is no history, rather than sending an empty context', async () => {
      const { service, complete } = makeService({ docs: [] });

      await service.ask('u1', 'what did I do');

      expect(complete.mock.calls[0][0].userMessage).toContain('no stored Cellix sessions');
    });
  });

  describe('read-only posture', () => {
    it('instructs the model that it cannot modify a workbook from this surface', async () => {
      const { service, complete } = makeService();
      await service.ask('u1', 'add a total row');

      expect(complete.mock.calls[0][0].systemPrompt).toContain('CANNOT modify any workbook');
    });

    it('returns no actions field at all — the web client has no Office.js host to apply one', async () => {
      const { service } = makeService();
      const answer = await service.ask('u1', 'add a total row');
      expect(answer).not.toHaveProperty('actions');
    });
  });
});
