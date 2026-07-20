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
import { SheetAnalyzerService } from './services/sheet-analyzer.service';
import { SmartDataQueryService } from './services/smart-data-query.service';
import { Tier0DirectService } from './services/tier0-direct.service';
import { Tier1SingleActionService } from './services/tier1-single-action.service';
import { Tier2GenerateVerifyService } from './services/tier2-generate-verify.service';

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
    Tier0DirectService,
    Tier1SingleActionService,
    Tier2GenerateVerifyService,
  ],
})
export class ExcelAiModule {}
