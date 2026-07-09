import { resolveConversationHistory } from '../src/excel-ai/utils/workbook-context-resolver.util';
import { ConversationRequestDto } from '../src/excel-ai/dto/conversation-request.dto';
import { ConversationMessageEntry } from '../src/excel-ai/schemas/conversation.schema';

function mongoMessage(
  id: string,
  role: 'user' | 'assistant',
  content: string,
): ConversationMessageEntry {
  return {
    id,
    role,
    content,
    type: role === 'user' ? 'command' : 'answer',
    timestamp: new Date('2026-01-01T00:00:00.000Z'),
  };
}

describe('resolveConversationHistory', () => {
  const baseRequest = {
    message: 'Follow up',
    sheetData: [],
    context: {},
  } as ConversationRequestDto;

  it('prefers explicit conversationHistory on the request', () => {
    const history = resolveConversationHistory(
      {
        ...baseRequest,
        conversationHistory: [{ role: 'user', content: 'from dto' }],
      },
      [mongoMessage('m1', 'assistant', 'mongo only')],
    );

    expect(history).toEqual([{ role: 'user', content: 'from dto' }]);
  });

  it('falls back to client previousMessages when dto history is empty', () => {
    const history = resolveConversationHistory(
      {
        ...baseRequest,
        context: {
          previousMessages: [
            {
              role: 'user',
              content: 'client msg',
              timestamp: '2026-01-01T00:00:00.000Z',
              type: 'command',
            },
          ],
        },
      },
      [mongoMessage('m1', 'assistant', 'mongo only')],
    );

    expect(history).toEqual([{ role: 'user', content: 'client msg' }]);
  });

  it('falls back to Mongo messages when client history is empty', () => {
    const history = resolveConversationHistory(
      {
        ...baseRequest,
        conversationId: 'conv_1',
        context: { previousMessages: [] },
      },
      [
        mongoMessage('m1', 'user', 'stored user'),
        mongoMessage('m2', 'assistant', 'stored assistant'),
      ],
    );

    expect(history).toEqual([
      { role: 'user', content: 'stored user' },
      { role: 'assistant', content: 'stored assistant' },
    ]);
  });

  it('returns empty history when no sources are available', () => {
    expect(resolveConversationHistory(baseRequest)).toEqual([]);
  });
});
