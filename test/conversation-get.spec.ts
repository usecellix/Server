import { GoneException, NotFoundException } from '@nestjs/common';
import { ConversationService } from '../src/excel-ai/services/conversation.service';

/**
 * getConversation reads only the Mongoose model — every other injected dependency is
 * unused here. Deriving the arity keeps this suite from breaking each time an
 * unrelated dependency is added to the constructor.
 */
function createService(model: { findOne: jest.Mock }) {
  type Deps = ConstructorParameters<typeof ConversationService>;
  const deps = [model, ...Array.from({ length: 32 }, () => ({}))] as unknown as Deps;
  return new ConversationService(...deps);
}

describe('ConversationService.getConversation', () => {
  it('returns conversation messages for an active thread', async () => {
    const updatedAt = new Date('2026-06-01T12:00:00.000Z');
    const findOne = jest.fn().mockReturnValue({
      lean: jest.fn().mockResolvedValue({
        conversationId: 'conv_1',
        messages: [{ id: 'm1', role: 'user', content: 'Find 2290' }],
        status: 'active',
        sheetSnapshot: { sheetName: 'Sheet1' },
        expiresAt: new Date(Date.now() + 60_000),
        updatedAt,
      }),
    });

    const service = createService({ findOne });
    const result = await service.getConversation('conv_1');

    expect(findOne).toHaveBeenCalledWith({ conversationId: 'conv_1' });
    expect(result).toEqual({
      conversationId: 'conv_1',
      messages: [{ id: 'm1', role: 'user', content: 'Find 2290' }],
      status: 'active',
      sheetSnapshot: { sheetName: 'Sheet1' },
      updatedAt,
    });
  });

  it('throws when conversation is missing', async () => {
    const findOne = jest.fn().mockReturnValue({
      lean: jest.fn().mockResolvedValue(null),
    });
    const service = createService({ findOne });

    await expect(service.getConversation('conv_missing')).rejects.toBeInstanceOf(
      NotFoundException,
    );
  });

  it('throws when conversation is expired', async () => {
    const findOne = jest.fn().mockReturnValue({
      lean: jest.fn().mockResolvedValue({
        conversationId: 'conv_expired',
        messages: [],
        status: 'active',
        expiresAt: new Date(Date.now() - 60_000),
      }),
    });
    const service = createService({ findOne });

    await expect(service.getConversation('conv_expired')).rejects.toBeInstanceOf(
      GoneException,
    );
  });
});
