import { BadRequestException, NotFoundException } from '@nestjs/common';
import { ConversationService } from '../src/excel-ai/services/conversation.service';

/**
 * TASKS.md #177 — rename/delete on top of #170/#171's user-scoped history.
 *
 * Same construction pattern as conversation-history.spec.ts: only the Mongoose
 * model is real, arity is derived so an unrelated new constructor dependency
 * doesn't break this suite.
 */
function createService(model: Record<string, jest.Mock>) {
  type Deps = ConstructorParameters<typeof ConversationService>;
  const deps = [model, ...Array.from({ length: 32 }, () => ({}))] as unknown as Deps;
  return new ConversationService(...deps);
}

describe('ConversationService.renameConversation (TASKS.md #177)', () => {
  function docWithOwner(userId: string | undefined) {
    return {
      conversationId: 'conv_1',
      userId,
      title: 'Old title',
      save: jest.fn().mockResolvedValue(undefined),
    };
  }

  it('sets a new title for the owner', async () => {
    const doc = docWithOwner('user_a');
    const service = createService({ findOne: jest.fn().mockResolvedValue(doc) });

    const result = await service.renameConversation('conv_1', 'user_a', 'Renamed chat');

    expect(doc.title).toBe('Renamed chat');
    expect(doc.save).toHaveBeenCalledTimes(1);
    expect(result).toEqual({ conversationId: 'conv_1', title: 'Renamed chat' });
  });

  it('trims and collapses whitespace the same way auto-derived titles are', async () => {
    const doc = docWithOwner('user_a');
    const service = createService({ findOne: jest.fn().mockResolvedValue(doc) });

    await service.renameConversation('conv_1', 'user_a', '  Multi   line\n\ntitle  ');

    expect(doc.title).toBe('Multi line title');
  });

  it('rejects an empty or whitespace-only title rather than saving a blank one', async () => {
    const doc = docWithOwner('user_a');
    const service = createService({ findOne: jest.fn().mockResolvedValue(doc) });

    await expect(service.renameConversation('conv_1', 'user_a', '   ')).rejects.toBeInstanceOf(
      BadRequestException,
    );
    expect(doc.save).not.toHaveBeenCalled();
  });

  it("rejects renaming another user's conversation as NOT_FOUND", async () => {
    const doc = docWithOwner('user_a');
    const service = createService({ findOne: jest.fn().mockResolvedValue(doc) });

    await expect(
      service.renameConversation('conv_1', 'user_b', 'Hijacked title'),
    ).rejects.toBeInstanceOf(NotFoundException);
    expect(doc.save).not.toHaveBeenCalled();
  });

  it('allows renaming an unowned pre-#170 conversation, matching read access', async () => {
    const doc = docWithOwner(undefined);
    const service = createService({ findOne: jest.fn().mockResolvedValue(doc) });

    const result = await service.renameConversation('conv_1', 'user_b', 'Claimed title');

    expect(result.title).toBe('Claimed title');
  });

  it('throws NOT_FOUND for a missing conversation', async () => {
    const service = createService({ findOne: jest.fn().mockResolvedValue(null) });

    await expect(
      service.renameConversation('conv_missing', 'user_a', 'New title'),
    ).rejects.toBeInstanceOf(NotFoundException);
  });
});

describe('ConversationService.deleteConversation (TASKS.md #177)', () => {
  function findOneChain(doc: unknown) {
    const lean = jest.fn().mockResolvedValue(doc);
    const select = jest.fn().mockReturnValue({ lean });
    const findOne = jest.fn().mockReturnValue({ select });
    return { findOne, select, lean };
  }

  it('deletes the conversation for its owner', async () => {
    const { findOne } = findOneChain({ userId: 'user_a' });
    const deleteOne = jest.fn().mockResolvedValue({ deletedCount: 1 });
    const service = createService({ findOne, deleteOne });

    await service.deleteConversation('conv_1', 'user_a');

    expect(deleteOne).toHaveBeenCalledWith({ conversationId: 'conv_1' });
  });

  it("refuses to delete another user's conversation and does not touch the collection", async () => {
    const { findOne } = findOneChain({ userId: 'user_a' });
    const deleteOne = jest.fn();
    const service = createService({ findOne, deleteOne });

    await expect(service.deleteConversation('conv_1', 'user_b')).rejects.toBeInstanceOf(
      NotFoundException,
    );
    expect(deleteOne).not.toHaveBeenCalled();
  });

  it('allows deleting an unowned pre-#170 conversation, matching read/rename access', async () => {
    const { findOne } = findOneChain({ userId: undefined });
    const deleteOne = jest.fn().mockResolvedValue({ deletedCount: 1 });
    const service = createService({ findOne, deleteOne });

    await service.deleteConversation('conv_1', 'user_b');

    expect(deleteOne).toHaveBeenCalledWith({ conversationId: 'conv_1' });
  });

  it('throws NOT_FOUND for a missing conversation rather than deleting nothing silently', async () => {
    const { findOne } = findOneChain(null);
    const deleteOne = jest.fn();
    const service = createService({ findOne, deleteOne });

    await expect(service.deleteConversation('conv_missing', 'user_a')).rejects.toBeInstanceOf(
      NotFoundException,
    );
    expect(deleteOne).not.toHaveBeenCalled();
  });
});
