import { BadRequestException, NotFoundException } from '@nestjs/common';
import { ConversationService } from '../src/excel-ai/services/conversation.service';
import {
  deriveConversationTitle,
  truncateTitle,
} from '../src/excel-ai/utils/conversation-title.util';

/**
 * TASKS.md #170/#171 — user-scoped conversation history.
 *
 * Only the Mongoose model is real; every other constructor dependency is unused
 * by the methods under test. Arity is derived rather than hardcoded, matching
 * conversation-get.spec.ts / conversation-workbook-id.spec.ts, so an unrelated
 * new dependency doesn't break this suite.
 */
function createService(model: Record<string, jest.Mock>) {
  type Deps = ConstructorParameters<typeof ConversationService>;
  const deps = [model, ...Array.from({ length: 32 }, () => ({}))] as unknown as Deps;
  return new ConversationService(...deps);
}

/** Mirrors the chained `.find().sort().limit().select().lean()` builder. */
function findChain(docs: unknown[]) {
  const lean = jest.fn().mockResolvedValue(docs);
  const select = jest.fn().mockReturnValue({ lean });
  const limit = jest.fn().mockReturnValue({ select });
  const sort = jest.fn().mockReturnValue({ limit });
  const find = jest.fn().mockReturnValue({ sort });
  return { find, sort, limit, select, lean };
}

function conversationDoc(overrides: Record<string, unknown> = {}) {
  return {
    conversationId: 'conv_1',
    userId: 'user_a',
    title: 'Sum the totals',
    status: 'active',
    messages: [
      { id: 'm1', role: 'user', content: 'Sum the totals' },
      { id: 'm2', role: 'assistant', content: 'Done — added a total row.' },
    ],
    updatedAt: new Date('2026-08-30T10:00:00.000Z'),
    ...overrides,
  };
}

describe('conversation title derivation (TASKS.md #171)', () => {
  it('uses the first user message as the title', () => {
    expect(
      deriveConversationTitle([
        { role: 'user', content: 'Add a Remarks column' },
        { role: 'assistant', content: 'Added.' },
        { role: 'user', content: 'Now bold the header' },
      ]),
    ).toBe('Add a Remarks column');
  });

  it('skips a leading assistant message rather than titling the chat with it', () => {
    expect(
      deriveConversationTitle([
        { role: 'assistant', content: 'How can I help?' },
        { role: 'user', content: 'Freeze the top row' },
      ]),
    ).toBe('Freeze the top row');
  });

  it('falls back to "New chat" for an empty or user-less conversation', () => {
    expect(deriveConversationTitle([])).toBe('New chat');
    expect(deriveConversationTitle([{ role: 'assistant', content: 'Hi' }])).toBe('New chat');
    expect(deriveConversationTitle([{ role: 'user', content: '   ' }])).toBe('New chat');
  });

  it('collapses newlines so a pasted multi-line prompt still renders as one row', () => {
    expect(truncateTitle('Add a column\n\nthen total it')).toBe('Add a column then total it');
  });

  it('truncates long titles with an ellipsis', () => {
    const long = 'a'.repeat(200);
    const result = truncateTitle(long);
    expect(result).toHaveLength(80);
    expect(result.endsWith('…')).toBe(true);
  });
});

describe('ConversationService.listConversations (TASKS.md #171)', () => {
  it('queries only the requesting user, newest first', async () => {
    const chain = findChain([conversationDoc()]);
    const service = createService({ find: chain.find });

    const result = await service.listConversations('user_a');

    expect(chain.find).toHaveBeenCalledWith({ userId: 'user_a' });
    expect(chain.sort).toHaveBeenCalledWith({ updatedAt: -1 });
    expect(result.conversations).toHaveLength(1);
    expect(result.conversations[0]).toMatchObject({
      conversationId: 'conv_1',
      title: 'Sum the totals',
      messageCount: 2,
    });
  });

  it('returns summaries only — never message bodies', async () => {
    const chain = findChain([conversationDoc()]);
    const service = createService({ find: chain.find });

    const result = await service.listConversations('user_a');

    expect(result.conversations[0]).not.toHaveProperty('messages');
    expect(result.conversations[0].firstMessage).toBe('Sum the totals');
    expect(result.conversations[0].lastMessage).toBe('Done — added a total row.');
  });

  it('derives a title for a pre-#171 doc stored without one', async () => {
    const chain = findChain([conversationDoc({ title: undefined })]);
    const service = createService({ find: chain.find });

    const result = await service.listConversations('user_a');

    expect(result.conversations[0].title).toBe('Sum the totals');
  });

  it('clamps limit to the page-size ceiling and never returns unbounded history', async () => {
    const chain = findChain([]);
    const service = createService({ find: chain.find });

    await service.listConversations('user_a', { limit: 5000 });

    // limit + 1 — the extra row is how "has another page" is answered.
    expect(chain.limit).toHaveBeenCalledWith(51);
  });

  it('applies a sane default limit when none is given', async () => {
    const chain = findChain([]);
    const service = createService({ find: chain.find });

    await service.listConversations('user_a');

    expect(chain.limit).toHaveBeenCalledWith(26);
  });

  it('emits a nextCursor only when a further page exists', async () => {
    const docs = Array.from({ length: 3 }, (_, index) =>
      conversationDoc({
        conversationId: `conv_${index}`,
        updatedAt: new Date(`2026-08-2${index}T10:00:00.000Z`),
      }),
    );
    const service = createService({ find: findChain(docs).find });

    const page = await service.listConversations('user_a', { limit: 2 });

    expect(page.conversations).toHaveLength(2);
    expect(page.nextCursor).toBe(docs[1].updatedAt.toISOString());
  });

  it('returns a null cursor on the last page', async () => {
    const service = createService({ find: findChain([conversationDoc()]).find });

    const page = await service.listConversations('user_a', { limit: 2 });

    expect(page.conversations).toHaveLength(1);
    expect(page.nextCursor).toBeNull();
  });

  it('translates a cursor into an updatedAt bound', async () => {
    const chain = findChain([]);
    const service = createService({ find: chain.find });

    await service.listConversations('user_a', { cursor: '2026-08-30T10:00:00.000Z' });

    expect(chain.find).toHaveBeenCalledWith({
      userId: 'user_a',
      updatedAt: { $lt: new Date('2026-08-30T10:00:00.000Z') },
    });
  });

  it('rejects an unparseable cursor rather than silently returning page one', async () => {
    const service = createService({ find: findChain([]).find });

    await expect(
      service.listConversations('user_a', { cursor: 'not-a-date' }),
    ).rejects.toBeInstanceOf(BadRequestException);
  });

  it('narrows by workbookId when one is supplied (open decision #173)', async () => {
    const chain = findChain([]);
    const service = createService({ find: chain.find });

    await service.listConversations('user_a', { workbookId: 'wb_1' });

    expect(chain.find).toHaveBeenCalledWith({ userId: 'user_a', workbookId: 'wb_1' });
  });
});

describe('ConversationService.getConversation — ownership (TASKS.md #171)', () => {
  function serviceWithDoc(doc: unknown) {
    return createService({
      findOne: jest.fn().mockReturnValue({ lean: jest.fn().mockResolvedValue(doc) }),
    });
  }

  it('returns the conversation to its owner', async () => {
    const service = serviceWithDoc(
      conversationDoc({ expiresAt: new Date(Date.now() + 60_000) }),
    );

    const result = await service.getConversation('conv_1', 'user_a');

    expect(result.conversationId).toBe('conv_1');
  });

  it("rejects another user's conversation as NOT_FOUND, not by returning it", async () => {
    const service = serviceWithDoc(
      conversationDoc({ expiresAt: new Date(Date.now() + 60_000) }),
    );

    await expect(service.getConversation('conv_1', 'user_b')).rejects.toBeInstanceOf(
      NotFoundException,
    );
  });

  it('still serves an unowned (pre-#170) conversation by id', async () => {
    const service = serviceWithDoc(
      conversationDoc({ userId: undefined, expiresAt: new Date(Date.now() + 60_000) }),
    );

    const result = await service.getConversation('conv_1', 'user_b');

    expect(result.conversationId).toBe('conv_1');
  });
});

describe('ConversationService.getOrCreateConversation — userId (TASKS.md #170)', () => {
  it('stamps the owner onto a newly created conversation', async () => {
    const create = jest.fn().mockImplementation((doc) => Promise.resolve(doc));
    const service = createService({ create, findOne: jest.fn() });

    const doc = await (service as any).getOrCreateConversation(undefined, undefined, 'user_a');

    expect(doc.userId).toBe('user_a');
  });

  it('omits userId entirely when there is no session (backward compatible)', async () => {
    const create = jest.fn().mockImplementation((doc) => Promise.resolve(doc));
    const service = createService({ create, findOne: jest.fn() });

    await (service as any).getOrCreateConversation(undefined, undefined, undefined);

    expect(create).toHaveBeenCalledWith(
      expect.not.objectContaining({ userId: expect.anything() }),
    );
  });

  it('backfills an owner onto a pre-#170 conversation being resumed', async () => {
    const existing = {
      conversationId: 'conv_old',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      userId: undefined as string | undefined,
      save: jest.fn().mockResolvedValue(undefined),
    };
    const service = createService({
      findOne: jest.fn().mockResolvedValue(existing),
      create: jest.fn(),
    });

    const doc = await (service as any).getOrCreateConversation('conv_old', undefined, 'user_a');

    expect(doc.userId).toBe('user_a');
    expect(existing.save).toHaveBeenCalledTimes(1);
  });

  it("refuses to continue another user's conversation", async () => {
    const existing = {
      conversationId: 'conv_theirs',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      userId: 'user_a',
      save: jest.fn().mockResolvedValue(undefined),
    };
    const service = createService({
      findOne: jest.fn().mockResolvedValue(existing),
      create: jest.fn(),
    });

    await expect(
      (service as any).getOrCreateConversation('conv_theirs', undefined, 'user_b'),
    ).rejects.toBeInstanceOf(NotFoundException);
    expect(existing.save).not.toHaveBeenCalled();
  });

  it('backfills workbookId and userId in a single save, not two', async () => {
    const existing = {
      conversationId: 'conv_old',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      userId: undefined as string | undefined,
      workbookId: undefined as string | undefined,
      save: jest.fn().mockResolvedValue(undefined),
    };
    const service = createService({
      findOne: jest.fn().mockResolvedValue(existing),
      create: jest.fn(),
    });

    await (service as any).getOrCreateConversation('conv_old', 'wb_1', 'user_a');

    expect(existing.save).toHaveBeenCalledTimes(1);
    expect(existing.userId).toBe('user_a');
    expect(existing.workbookId).toBe('wb_1');
  });
});

describe('Conversation retention (TASKS.md #170)', () => {
  it('defaults to a 90-day window rather than the old 24h scratch buffer', async () => {
    // Imported lazily so the module-level env read happens with the default.
    const { CONVERSATION_TTL_MS } = await import(
      '../src/excel-ai/schemas/conversation.schema'
    );
    expect(CONVERSATION_TTL_MS).toBe(90 * 24 * 60 * 60 * 1000);
  });
});
