/**
 * Multi-sheet / large-write confirmation resume.
 * When the assistant only offers work ("want me to apply?") without actions,
 * store a pendingWritePlan so short affirmations ("yes do it") re-enter the write path
 * instead of the local "I see N rows…" dead end.
 */
import { ConversationMessageEntry } from '../schemas/conversation.schema';

export interface PendingWritePlan {
  /** Original user request that triggered the confirm offer. */
  originalPrompt: string;
  offeredAt?: string;
  /** Short human summary from the assistant message (optional). */
  summary?: string;
}

const AFFIRMATION_RE =
  /^(yes|yep|yeah|yup|ok|okay|sure|go ahead|proceed|do it|yes do it|please do|please proceed|apply|apply it|apply that|confirm|sounds good|that works|let'?s do it|go for it)(\s*[.!]*)?$/i;

const AFFIRMATION_PREFIX_RE =
  /^(yes|yep|yeah|sure|ok|okay|go ahead)\s*[,.]?\s*(do it|please|apply|proceed|go ahead)?\s*[.!]?$/i;

/** User message is a short accept of a prior offer. */
export function isAffirmationMessage(message: string): boolean {
  const t = message.trim();
  if (!t || t.length > 80) return false;
  if (AFFIRMATION_RE.test(t)) return true;
  if (AFFIRMATION_PREFIX_RE.test(t)) return true;
  return /^(yes|ok|sure|go ahead)\b/i.test(t) && t.split(/\s+/).length <= 6;
}

/**
 * Assistant prose offered to create/apply a change without emitting actions.
 */
export function isConfirmationOfferText(text: string): boolean {
  if (!text?.trim()) return false;
  const lower = text.toLowerCase();
  return (
    /\bwant me to apply\b/.test(lower) ||
    /\bneed your confirmation\b/.test(lower) ||
    /\bwould need your confirmation\b/.test(lower) ||
    /\bshall i (create|build|apply|add|proceed)\b/.test(lower) ||
    /\bshould i (create|build|apply|add|proceed)\b/.test(lower) ||
    /\bi can create\b/.test(lower) ||
    /\blet me know if you (want|meant)\b/.test(lower) ||
    /\breply (yes|confirm)\b/.test(lower) ||
    /\bconfirm to (add|create|apply)\b/.test(lower) ||
    /\bif that works,?\s*(i('ll| will)|say yes)\b/.test(lower)
  );
}

export function lastUserPromptBeforeAssistant(
  history: ConversationMessageEntry[],
  assistantIndex: number,
): string | undefined {
  for (let i = assistantIndex - 1; i >= 0; i -= 1) {
    if (history[i]?.role === 'user' && history[i]!.content?.trim()) {
      return history[i]!.content.trim();
    }
  }
  return undefined;
}

/**
 * Find the latest pending write plan from conversation history.
 * Prefers explicit metadata; falls back to confirmation offer + prior user text.
 */
export function findPendingWritePlan(
  history: ConversationMessageEntry[],
): PendingWritePlan | null {
  for (let i = history.length - 1; i >= 0; i -= 1) {
    const entry = history[i]!;
    if (entry.role !== 'assistant') continue;

    const metaPlan = entry.metadata?.pendingWritePlan as PendingWritePlan | undefined;
    if (metaPlan?.originalPrompt?.trim()) {
      return {
        originalPrompt: metaPlan.originalPrompt.trim(),
        offeredAt: metaPlan.offeredAt,
        summary: metaPlan.summary,
      };
    }

    if (isConfirmationOfferText(entry.content)) {
      const prior = lastUserPromptBeforeAssistant(history, i);
      if (prior && prior.length >= 20 && !isAffirmationMessage(prior)) {
        return {
          originalPrompt: prior,
          summary: entry.content.slice(0, 240),
          offeredAt: entry.timestamp ? new Date(entry.timestamp).toISOString() : undefined,
        };
      }
    }
  }
  return null;
}

/** Prompt injected when user affirms a prior offer. */
export function buildResumedWritePrompt(plan: PendingWritePlan): string {
  return (
    `User confirmed: proceed with this previously planned write. Do not ask for confirmation again. ` +
    `Implement fully with real sheet actions (CREATE_SHEET / WRITE_TABLE / FORMAT_RANGE / SET_FORMULA as needed). ` +
    `Do not overwrite existing unrelated data (e.g. Purchase Register) unless the request explicitly requires it — prefer adding new sheets. ` +
    `Original request:\n${plan.originalPrompt}`
  );
}

export function shouldStorePendingWritePlan(
  answer: string,
  hasActions: boolean,
): boolean {
  if (hasActions) return false;
  return isConfirmationOfferText(answer);
}

export function buildPendingWritePlanMetadata(
  originalPrompt: string,
  answer: string,
): { pendingWritePlan: PendingWritePlan; pendingIntent: string } {
  return {
    pendingIntent: 'write_confirm',
    pendingWritePlan: {
      originalPrompt: originalPrompt.trim(),
      offeredAt: new Date().toISOString(),
      summary: answer.slice(0, 320),
    },
  };
}

/** Honest copy when local engine cannot execute a write affirmation. */
export function localWriteUnavailableMessage(plan: PendingWritePlan | null): string {
  if (plan?.originalPrompt) {
    return (
      `I still have your pending request ready: **${clip(plan.originalPrompt, 160)}**. ` +
      `I can't create the sheets in limited local mode (AI provider unavailable). ` +
      `Restore OPENROUTER_API_KEY, restart the backend, then send **yes** again — or restate the full request in Action mode.`
    );
  }
  return (
    `I don't have a pending change to apply, and I can't invent a multi-sheet write without the AI write path. ` +
    `Restate what you want built (e.g. monthly sheets + main dashboard), with OPENROUTER_API_KEY configured, in Action mode.`
  );
}

export function localActionWithoutLlmMessage(): string {
  return (
    `I recognized a write request, but the AI write path is unavailable right now, so I won't guess structural changes. ` +
    `Set OPENROUTER_API_KEY in the backend .env and retry — or be more specific about a simple local action (add row, sum column).`
  );
}

function clip(s: string, n: number): string {
  const t = s.replace(/\s+/g, ' ').trim();
  return t.length <= n ? t : `${t.slice(0, n - 1)}…`;
}
