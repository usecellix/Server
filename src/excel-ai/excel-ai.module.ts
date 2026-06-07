import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { ConversationController } from './conversation.controller';
import { Conversation, ConversationSchema } from './schemas/conversation.schema';
import { ConversationEngineService } from './services/conversation-engine.service';
import { ConversationService } from './services/conversation.service';
import { DataQueryService } from './services/data-query.service';
import { IntentClassifierService } from './services/intent-classifier.service';
import { OpenRouterService } from './services/openrouter.service';
import { SheetAnalyzerService } from './services/sheet-analyzer.service';

@Module({
  imports: [
    AppConfigModule,
    MongooseModule.forFeature([{ name: Conversation.name, schema: ConversationSchema }]),
  ],
  controllers: [ConversationController],
  providers: [
    ConversationService,
    ConversationEngineService,
    OpenRouterService,
    SheetAnalyzerService,
    IntentClassifierService,
    DataQueryService,
  ],
})
export class ExcelAiModule {}
