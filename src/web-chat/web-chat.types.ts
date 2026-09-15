/**
 * Types for the web chat surface (`web/` Next.js app).
 *
 * The web app is an ASK-MODE-ONLY companion over the user's Excel add-in
 * history — it never mutates a workbook, so nothing here carries a
 * SheetAction. That is not a temporary limitation: the web client has no
 * Office.js host to apply an action to, so a write-shaped answer would be
 * unapplicable by construction. `WebChatAnswer` therefore has no `actions`
 * field at all, rather than an always-empty one.
 */

/** How a web message was priced. Mirrors the existing ask-mode Q&A tiers. */
export type WebChatComplexity = 'simple' | 'complex';

export interface WebChatCitation {
  /** The Excel conversation this claim was drawn from. */
  conversationId: string;
  /** That conversation's title at answer time, for display without a re-fetch. */
  title: string;
}

export interface WebChatAnswer {
  answer: string;
  /**
   * Which stored Excel conversations informed the answer. Present so the UI
   * can show "based on 2 of your sessions" rather than presenting a synthesis
   * over the user's own history as if it were free-standing knowledge.
   */
  citations: WebChatCitation[];
  complexity: WebChatComplexity;
  creditsDeducted: number;
  /** Balance AFTER the debit — what the sidebar should render immediately. */
  newBalance: number;
}
