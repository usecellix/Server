// cellix_backend/src/excel-ai/prompts/chitchat-prompt.ts

/**
 * System prompt for LlmRouterService.classifyIntent — a LOW-tier pre-check that
 * runs before tier classification / router.route(). Single-purpose: return only
 * the label, nothing else, so the caller can parse the raw response directly
 * (no JSON, unlike ROUTER_SYSTEM_PROMPT).
 */
export const CHITCHAT_CLASSIFIER_SYSTEM_PROMPT = `Classify the user's message into exactly one category. Respond with only the label, nothing else.
CHITCHAT — greetings, small talk, thanks, goodbyes, or questions about what the assistant is / can do.
TASK — anything referring to the spreadsheet, data, formulas, formatting, or asking the assistant to look at, analyze, or change something.`;

export function buildChitchatClassifierUserMessage(userMessage: string): string {
  return `Message: "${userMessage}"\nLabel:`;
}

/**
 * Persona system prompt for ChitchatService — the ONE-call reply for messages
 * classifyIntent labeled CHITCHAT. Never proposes actions: there is no
 * workbook context loaded on this path, so there is nothing to ground an
 * action in.
 */
export const CHITCHAT_PERSONA_SYSTEM_PROMPT = `You are Cellix, an AI agent that works inside Excel. Casual greetings/small talk: respond warmly and briefly, then ask what they need help with. 'Who are you' style questions: introduce yourself as Cellix in one sentence — an agent that reads, edits, and verifies changes in their spreadsheet — then ask what they need. Never propose actions or mention tools for these messages.`;
