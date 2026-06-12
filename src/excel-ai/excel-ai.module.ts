import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AuditModule } from '../audit/audit.module';
import { AppConfigModule } from '../config/app-config.module';
import { ConversationController } from './conversation.controller';
import { Conversation, ConversationSchema } from './schemas/conversation.schema';
import { ConversationEngineService } from './services/conversation-engine.service';
import { ConversationService } from './services/conversation.service';
import { DataQueryService } from './services/data-query.service';
import { IntentClassifierService } from './services/intent-classifier.service';
import { ModelRouter } from './llm/model-router';
import { OpenRouterService } from './services/openrouter.service';
import { SheetAnalyzerService } from './services/sheet-analyzer.service';

@Module({
  imports: [
    AppConfigModule,
    AuditModule,
    MongooseModule.forFeature([{ name: Conversation.name, schema: ConversationSchema }]),
  ],
  controllers: [ConversationController],
  providers: [
    ConversationService,
    ConversationEngineService,
    ModelRouter,
    OpenRouterService,
    SheetAnalyzerService,
    IntentClassifierService,
    DataQueryService,
  ],
})
export class ExcelAiModule {}
