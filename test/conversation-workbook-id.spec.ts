import { ConversationService } from '../src/excel-ai/services/conversation.service';

/**
 * getOrCreateConversation (private) is exercised directly via `as any`, matching
 * the established pattern for testing private ConversationService methods
 * (see mode-plan-only.spec.ts / mode-selector.e2e.spec.ts's streamPlanOnly calls).
 * Only the model is real; every other constructor dependency is unused by this
 * method, mirroring conversation-get.spec.ts's arity-tolerant helper.
 */
function createService(model: { findOne: jest.Mock; create: jest.Mock }) {
  type Deps = ConstructorParameters<typeof ConversationService>;
  const deps = [model, ...Array.from({ length: 32 }, () => ({}))] as unknown as Deps;
  return new ConversationService(...deps);
}

describe('ConversationService.getOrCreateConversation — workbookId (TASKS.md #23)', () => {
  it('persists workbookId on a newly created conversation', async () => {
    const create = jest.fn().mockImplementation((doc) => Promise.resolve(doc));
    const findOne = jest.fn();
    const service = createService({ findOne, create });

    const doc = await (service as any).getOrCreateConversation(undefined, 'wb_shared-1');

    expect(create).toHaveBeenCalledWith(
      expect.objectContaining({ workbookId: 'wb_shared-1' }),
    );
    expect(doc.workbookId).toBe('wb_shared-1');
  });

  it('two separate (new) conversations started with the same workbookId both persist it', async () => {
    const create = jest.fn().mockImplementation((doc) => Promise.resolve(doc));
    const findOne = jest.fn();
    const service = createService({ findOne, create });

    const first = await (service as any).getOrCreateConversation(undefined, 'wb_shared-2');
    const second = await (service as any).getOrCreateConversation(undefined, 'wb_shared-2');

    expect(first.conversationId).not.toBe(second.conversationId);
    expect(first.workbookId).toBe('wb_shared-2');
    expect(second.workbookId).toBe('wb_shared-2');
  });

  it('a conversation started without a workbookId behaves exactly as before (backward compatible)', async () => {
    const create = jest.fn().mockImplementation((doc) => Promise.resolve(doc));
    const findOne = jest.fn();
    const service = createService({ findOne, create });

    const doc = await (service as any).getOrCreateConversation(undefined, undefined);

    expect(create).toHaveBeenCalledWith(
      expect.not.objectContaining({ workbookId: expect.anything() }),
    );
    expect(doc.workbookId).toBeUndefined();
  });

  it('backfills workbookId onto an existing conversation that has none yet', async () => {
    const existing = {
      conversationId: 'conv_existing',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      workbookId: undefined as string | undefined,
      save: jest.fn().mockResolvedValue(undefined),
    };
    const findOne = jest.fn().mockResolvedValue(existing);
    const create = jest.fn();
    const service = createService({ findOne, create });

    const doc = await (service as any).getOrCreateConversation('conv_existing', 'wb_late-mint');

    expect(doc.workbookId).toBe('wb_late-mint');
    expect(existing.save).toHaveBeenCalledTimes(1);
  });

  it('never overwrites an existing workbookId on an already-tagged conversation', async () => {
    const existing = {
      conversationId: 'conv_existing',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      workbookId: 'wb_original',
      save: jest.fn().mockResolvedValue(undefined),
    };
    const findOne = jest.fn().mockResolvedValue(existing);
    const create = jest.fn();
    const service = createService({ findOne, create });

    const doc = await (service as any).getOrCreateConversation('conv_existing', 'wb_different');

    expect(doc.workbookId).toBe('wb_original');
    expect(existing.save).not.toHaveBeenCalled();
  });

  it('resuming an existing conversation without a workbookId leaves it untouched', async () => {
    const existing = {
      conversationId: 'conv_existing',
      messages: [],
      expiresAt: new Date(Date.now() + 60_000),
      workbookId: undefined as string | undefined,
      save: jest.fn().mockResolvedValue(undefined),
    };
    const findOne = jest.fn().mockResolvedValue(existing);
    const create = jest.fn();
    const service = createService({ findOne, create });

    const doc = await (service as any).getOrCreateConversation('conv_existing', undefined);

    expect(doc.workbookId).toBeUndefined();
    expect(existing.save).not.toHaveBeenCalled();
  });
});
