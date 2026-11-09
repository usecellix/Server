import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { HydratedDocument } from 'mongoose';

export type SheetsLinkKeyDocument = HydratedDocument<SheetsLinkKey>;

/** One Google Sheets link key. Only the SHA-256 of the key is stored — never the key. */
@Schema({ collection: 'sheets_link_keys', versionKey: false })
export class SheetsLinkKey {
  @Prop({ type: String, required: true, unique: true, index: true })
  keyHash!: string;

  @Prop({ type: String, required: true, index: true })
  userId!: string;

  @Prop({ type: String })
  email?: string | null;

  @Prop({ type: String })
  name?: string | null;

  @Prop({ type: Date, default: () => new Date() })
  createdAt!: Date;

  @Prop({ type: Date })
  lastUsedAt?: Date;

  @Prop({ type: Date, required: true })
  expiresAt!: Date;
}

export const SheetsLinkKeySchema = SchemaFactory.createForClass(SheetsLinkKey);
SheetsLinkKeySchema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });
