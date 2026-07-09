import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AgentsModule } from '../agents/agents.module';
import { AuditModule } from '../audit/audit.module';
import { AppConfigModule } from '../config/app-config.module';
import { FormulaModule } from '../formula/formula.module';
import { LlmModule } from '../llm/llm.module';
import { ConversationController } from './conversation.controller';
import { Conversation, ConversationSchema } from './schemas/conversation.schema';
import { ConversationEngineService } from './services/conversation-engine.service';
import { ContextCacheService } from './services/context-cache.service';
import { ConversationService } from './services/conversation.service';
import { FindExportService } from './services/find-export.service';
import { DataQueryService } from './services/data-query.service';
import { IntentClassifierService } from './services/intent-classifier.service';
import { LlmRouterService } from './services/llm-router.service';
import { SmartDataQueryService } from './services/smart-data-query.service';

@Module({
  imports: [
    AppConfigModule,
    AuditModule,
    LlmModule,
    FormulaModule,
    AgentsModule,
    MongooseModule.forFeature([{ name: Conversation.name, schema: ConversationSchema }]),
  ],
  controllers: [ConversationController],
  providers: [
    ConversationService,
    ConversationEngineService,
    ContextCacheService,
    SheetAnalyzerService,
    IntentClassifierService,
    LlmRouterService,
    DataQueryService,
    FindExportService,
    SmartDataQueryService,
  ],
})
export class ExcelAiModule {}
