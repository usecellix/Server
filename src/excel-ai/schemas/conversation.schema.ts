import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { Document } from 'mongoose';
import type { TurnActionRecord } from '../utils/turn-action-history.util';

export type ConversationDocument = Conversation & Document;

/**
 * Conversation retention window (TASKS.md #170) — 90 days, overridable via
 * `CONVERSATION_TTL_HOURS`.
 *
 * Defined here rather than in `conversation.service.ts` because both the schema
 * default and the service's per-save refresh must agree. They previously did not:
 * the schema declared 24h while the service wrote `Date.now() + 168h` on create,
 * so the schema default was dead code and the real retention was 7 days. One
 * exported constant, imported by both, removes the chance of that drifting again.
 */
export const CONVERSATION_TTL_MS =
  Number(process.env.CONVERSATION_TTL_HOURS ?? 24 * 90) * 60 * 60 * 1000;

@Schema({ _id: false })
export class ConversationMessageEntry {
  @Prop({ required: true })
  id!: string;

  @Prop({ required: true, enum: ['user', 'assistant'] })
  role!: 'user' | 'assistant';

  @Prop({ required: true })
  content!: string;

  @Prop({
    enum: ['question', 'answer', 'command', 'clarification'],
    default: 'command',
  })
  type!: 'question' | 'answer' | 'command' | 'clarification';

  @Prop({ default: () => new Date() })
  timestamp!: Date;

  @Prop({ type: Object })
  metadata?: {
    actions?: unknown[];
    changeSetId?: string;
    questionOptions?: string[];
    pendingIntent?: string;
    /**
     * Multi-sheet / large write offer awaiting short affirmation ("yes do it").
     * Rehydrated on the next turn to force the write path.
     */
    pendingWritePlan?: {
      originalPrompt: string;
      offeredAt?: string;
      summary?: string;
    };
    ambiguityScore?: number;
    /** Spec 12 — successful early subtasks delivered when a later step fails */
    partialProgress?: boolean;
    failedSubtask?: { subtaskId: string; reason: string } | null;
    /** Spec 18 — structured chart/table range identity for follow-up turns */
    turnActionRecords?: TurnActionRecord[];
  };
}

export const ConversationMessageEntrySchema = SchemaFactory.createForClass(ConversationMessageEntry);

@Schema({
  timestamps: true,
  collection: 'conversations',
})
export class Conversation {
  @Prop({ required: true, unique: true, index: true })
  conversationId!: string;

  /**
   * Durable per-workbook identity (TASKS.md #22-23, ARCHITECTURE.md AD-9). Optional
   * and additive — minted client-side via Office.js document.settings, so it
   * survives past this conversation's 24h TTL, unlike `conversationId` itself.
   * A conversation created before this field existed simply has none; nothing
   * backfills it retroactively (DATABASE_SCHEMA.md §6.1/§7).
   */
  @Prop({ type: String, required: false, index: true })
  workbookId?: string;

  /**
   * Owner of this conversation (TASKS.md #170). Resolved server-side from the
   * authenticated session — never read from the request body, since that would
   * let a caller claim someone else's history.
   *
   * Optional and additive, same discipline as `workbookId` above: a conversation
   * written before this field existed simply has none, and nothing backfills it.
   * Those orphans are invisible to the history list (#171 filters on `userId`),
   * which is the correct failure mode — better to under-report than to leak an
   * unowned conversation into some arbitrary user's list.
   */
  @Prop({ type: String, required: false, index: true })
  userId?: string;

  /**
   * First user message, truncated — the history list's label (TASKS.md #171).
   * Denormalized onto the doc so listing conversations never has to load and
   * scan `messages`, which is the whole reason the list response stays small.
   */
  @Prop({ type: String, required: false })
  title?: string;

  @Prop({ type: [ConversationMessageEntrySchema], default: [] })
  messages!: ConversationMessageEntry[];

  @Prop({ type: Object })
  sheetSnapshot?: {
    rowCount: number;
    columnCount: number;
    headers: string[];
  };

  /**
   * Hash of the last TOON-compressed payload.
   * Used by ContextCacheService to skip re-analysis when sheet hasn't changed.
   */
  @Prop({ type: String, required: false })
  lastSheetHash?: string;

  /**
   * The cached promptContext from the last successful SheetAnalyzer run.
   * Only valid when lastSheetHash matches the current turn's TOON hash.
   */
  @Prop({ type: String, required: false })
  cachedPromptContext?: string;

  @Prop({ enum: ['active', 'completed', 'error'], default: 'active' })
  status!: 'active' | 'completed' | 'error';

  /**
   * Retention horizon (TASKS.md #170). Was 24h, which made conversations a
   * scratch buffer for resuming *today's* thread rather than history a user
   * could come back to. Now 90 days — deliberately time-boxed rather than
   * indefinite, matching the change-set/audit precedent of "durable but not
   * unbounded" (ARCHITECTURE.md AD-4 flags unTTL'd audit collections as a gap
   * worth not repeating by accident).
   *
   * Refreshed on every saved message, so the window is 90 days of inactivity,
   * not 90 days from creation.
   */
  @Prop({ default: () => new Date(Date.now() + CONVERSATION_TTL_MS) })
  expiresAt!: Date;
}

export const ConversationSchema = SchemaFactory.createForClass(Conversation);
ConversationSchema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });
/** Backs #171's `find({ userId }).sort({ updatedAt: -1 })` history listing. */
ConversationSchema.index({ userId: 1, updatedAt: -1 });
