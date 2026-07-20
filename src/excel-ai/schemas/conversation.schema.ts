import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { Document } from 'mongoose';

export type ConversationDocument = Conversation & Document;

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
    ambiguityScore?: number;
    /** Spec 12 — successful early subtasks delivered when a later step fails */
    partialProgress?: boolean;
    failedSubtask?: { subtaskId: string; reason: string } | null;
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

  @Prop({ default: () => new Date(Date.now() + 24 * 60 * 60 * 1000) })
  expiresAt!: Date;
}

export const ConversationSchema = SchemaFactory.createForClass(Conversation);
ConversationSchema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });
