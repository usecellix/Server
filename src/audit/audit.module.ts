import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AuditController } from './audit.controller';
import { AuditService } from './audit.service';
import { AuditEntry, AuditEntrySchema } from './schemas/audit-entry.schema';
import { AuditLog, AuditLogSchema } from './schemas/audit-log.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      { name: AuditEntry.name, schema: AuditEntrySchema },
      { name: AuditLog.name, schema: AuditLogSchema },
    ]),
  ],
  providers: [AuditService],
  controllers: [AuditController],
  exports: [AuditService],
})
export class AuditModule {}
