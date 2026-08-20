import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import {
  WorkflowTrace,
  WorkflowTraceSchema,
} from '../common/logging/schemas/workflow-trace.schema';
import { LoggingModule } from '../common/logging/logging.module';
import { Conversation, ConversationSchema } from '../excel-ai/schemas/conversation.schema';
import { AuditController } from './audit.controller';
import { AuditService } from './audit.service';
import { CheckpointController } from './checkpoint.controller';
import { CheckpointService } from './checkpoint.service';
import { ChangeSetController } from './change-set.controller';
import { ChangeSetService } from './change-set.service';
import { AuditEntry, AuditEntrySchema } from './schemas/audit-entry.schema';
import { AuditLog, AuditLogSchema } from './schemas/audit-log.schema';
import { ChangeSet, ChangeSetSchema } from './schemas/change-set.schema';
import { Checkpoint, CheckpointSchema } from './schemas/checkpoint.schema';
import { TierAMetricsService } from './tier-a-metrics.service';

@Module({
  imports: [
    LoggingModule,
    MongooseModule.forFeature([
      { name: AuditEntry.name, schema: AuditEntrySchema },
      { name: AuditLog.name, schema: AuditLogSchema },
      { name: ChangeSet.name, schema: ChangeSetSchema },
      { name: Checkpoint.name, schema: CheckpointSchema },
      // Read-only, resolve-workbookId-by-conversationId lookup only (TASKS.md #24).
      // AuditModule does not own or write this collection — ExcelAiModule does.
      { name: Conversation.name, schema: ConversationSchema },
      // Read-only, route/tier/verifier-outcome lookup for Tier A metrics (TASKS.md #48-51).
      // AuditModule does not own or write this collection — LoggingModule does.
      { name: WorkflowTrace.name, schema: WorkflowTraceSchema },
    ]),
  ],
  providers: [AuditService, ChangeSetService, CheckpointService, TierAMetricsService],
  controllers: [AuditController, ChangeSetController, CheckpointController],
  exports: [AuditService, ChangeSetService, CheckpointService, TierAMetricsService],
})
export class AuditModule {}
