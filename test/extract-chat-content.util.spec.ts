import { extractChatContent } from '../src/excel-ai/utils/extract-chat-content.util';

describe('extractChatContent', () => {
  it('returns string content as-is', () => {
    expect(extractChatContent('{"a":1}')).toBe('{"a":1}');
  });

  it('concatenates text parts from content arrays', () => {
    expect(
      extractChatContent([
        { type: 'text', text: '{"subtasks":' },
        { type: 'text', text: '[]}' },
      ]),
    ).toBe('{"subtasks":[]}');
  });

  it('handles null and undefined', () => {
    expect(extractChatContent(null)).toBe('');
    expect(extractChatContent(undefined)).toBe('');
  });

  it('returns empty string for unsupported shapes', () => {
    expect(extractChatContent(42)).toBe('');
    expect(extractChatContent({})).toBe('');
  });
});
