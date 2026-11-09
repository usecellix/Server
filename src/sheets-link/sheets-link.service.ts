import { Injectable, OnModuleInit } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { SheetsLinkKey, SheetsLinkKeyDocument } from './sheets-link.schema';
import { hashSheetsKey, LinkedUser, newSheetsKey, SHEETS_KEY_TTL_DAYS, SheetsKeyResolver } from './sheets-link.key';

@Injectable()
export class SheetsLinkService implements OnModuleInit {
  constructor(@InjectModel(SheetsLinkKey.name) private readonly keys: Model<SheetsLinkKeyDocument>) {}

  onModuleInit(): void {
    SheetsKeyResolver.resolve = (keyHash) => this.resolve(keyHash);
  }

  /** Creates a key for this user. The plain key is returned ONCE and never stored. */
  async create(user: LinkedUser): Promise<{ key: string; expiresAt: Date }> {
    const key = newSheetsKey();
    const expiresAt = new Date(Date.now() + SHEETS_KEY_TTL_DAYS * 86_400_000);
    await this.keys.create({ keyHash: hashSheetsKey(key), userId: user.userId, email: user.email, name: user.name, expiresAt });
    return { key, expiresAt };
  }

  async revokeAll(userId: string): Promise<number> {
    const res = await this.keys.deleteMany({ userId });
    return res.deletedCount ?? 0;
  }

  async resolve(keyHash: string): Promise<LinkedUser | null> {
    const doc = await this.keys
      .findOneAndUpdate({ keyHash, expiresAt: { $gt: new Date() } }, { $set: { lastUsedAt: new Date() } }, { new: true })
      .lean()
      .exec();
    return doc ? { userId: doc.userId, email: doc.email ?? null, name: doc.name ?? null } : null;
  }
}
