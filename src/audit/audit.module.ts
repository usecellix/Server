import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AuditService } from './audit.service';
import { AuditEntry, AuditEntrySchema } from './schemas/audit-entry.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      {
        name: AuditEntry.name,
        schema: AuditEntrySchema,
      },
    ]),
  ],
  providers: [AuditService],
  exports: [AuditService],
})
export class AuditModule {}
