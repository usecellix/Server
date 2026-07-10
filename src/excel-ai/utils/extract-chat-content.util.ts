/**
 * Normalizes assistant message content from OpenRouter chat completions.
 * Content may be a plain string or an array of typed parts ({ type: 'text', text: '...' }).
 */
export function extractChatContent(content: unknown): string {
  if (content == null) {
    return '';
  }
  if (typeof content === 'string') {
    return content;
  }
  if (!Array.isArray(content)) {
    return '';
  }

  const parts: string[] = [];
  for (const item of content) {
    if (typeof item === 'string') {
      parts.push(item);
      continue;
    }
    if (!item || typeof item !== 'object') {
      continue;
    }
    const record = item as Record<string, unknown>;
    if (typeof record.text === 'string') {
      parts.push(record.text);
    } else if (typeof record.content === 'string') {
      parts.push(record.content);
    }
  }
  return parts.join('');
}
