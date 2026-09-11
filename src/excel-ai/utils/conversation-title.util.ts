import type { ConversationMessageEntry } from '../schemas/conversation.schema';

/** Longest stored/returned conversation title — a list label, not a preview pane. */
export const TITLE_MAX_LENGTH = 80;

/**
 * Shorten to `TITLE_MAX_LENGTH`, collapsing whitespace first.
 *
 * Newlines matter here: a pasted multi-line prompt would otherwise render as a
 * label with hard breaks in it, which breaks the single-line list row.
 */
export function truncateTitle(raw: string, maxLength = TITLE_MAX_LENGTH): string {
  const collapsed = raw.replace(/\s+/g, ' ').trim();
  if (collapsed.length <= maxLength) return collapsed;
  return `${collapsed.slice(0, maxLength - 1).trimEnd()}…`;
}

/**
 * Label for a conversation in the history list (TASKS.md #171) — the first user
 * message, truncated.
 *
 * Falls back to 'New chat' rather than an empty string so a conversation created
 * but never sent still renders as a row a user can click, instead of a blank one
 * that looks like a bug.
 */
export function deriveConversationTitle(
  messages: Pick<ConversationMessageEntry, 'role' | 'content'>[],
): string {
  const first = messages.find((message) => message.role === 'user' && message.content?.trim());
  if (!first) return 'New chat';
  return truncateTitle(first.content) || 'New chat';
}
