import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AuditController } from './audit.controller';
import { AuditService } from './audit.service';
import { ChangeSetController } from './change-set.controller';
import { ChangeSetService } from './change-set.service';
import { AuditEntry, AuditEntrySchema } from './schemas/audit-entry.schema';
import { AuditLog, AuditLogSchema } from './schemas/audit-log.schema';
import { ChangeSet, ChangeSetSchema } from './schemas/change-set.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      { name: AuditEntry.name, schema: AuditEntrySchema },
      { name: AuditLog.name, schema: AuditLogSchema },
      { name: ChangeSet.name, schema: ChangeSetSchema },
    ]),
  ],
  providers: [AuditService, ChangeSetService],
  controllers: [AuditController, ChangeSetController],
  exports: [AuditService, ChangeSetService],
})
export class AuditModule {}
