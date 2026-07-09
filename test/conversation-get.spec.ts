import { GoneException, NotFoundException } from '@nestjs/common';
import { ConversationService } from '../src/excel-ai/services/conversation.service';

function createService(model: { findOne: jest.Mock }) {
  return new ConversationService(
    model as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
    {} as never,
  );
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
