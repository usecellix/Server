import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { AuditEntry, AuditEntryDocument } from './schemas/audit-entry.schema';

export interface CreateAuditEntryInput {
  requestId: string;
  processName: string;
  action: string;
  userId?: string;
  confidence?: number;
  payload?: unknown;
  result?: unknown;
}

@Injectable()
export class AuditService {
  constructor(
    @InjectModel(AuditEntry.name)
    private readonly auditEntryModel: Model<AuditEntryDocument>,
  ) {}

  async create(input: CreateAuditEntryInput): Promise<AuditEntry> {
    return this.auditEntryModel.create(input);
  }

  async findByRequestId(requestId: string): Promise<AuditEntry[]> {
    return this.auditEntryModel.find({ requestId }).sort({ createdAt: 1 }).lean<AuditEntry[]>().exec();
  }
}
