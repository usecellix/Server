import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { CreditModule } from '../credit/credit.module';
import { LlmModule } from '../llm/llm.module';
import { Conversation, ConversationSchema } from '../excel-ai/schemas/conversation.schema';
import { WebChatController } from './web-chat.controller';
import { WebChatService } from './web-chat.service';

/**
 * Ask-mode web chat over stored Excel sessions (`web/` Next.js app).
 *
 * Imports CreditModule rather than re-implementing a balance check: the web
 * surface spends from the SAME credit account as the Excel add-in, which is
 * the whole point of reusing this backend instead of standing up a separate
 * one. Registering the Conversation model here (rather than importing
 * ExcelAiModule wholesale) keeps this module's dependency surface to the one
 * collection it actually reads — it has no business pulling in the Tier 0-3
 * write pipeline it can never invoke.
 */
@Module({
  imports: [
    AppConfigModule,
    CreditModule,
    LlmModule,
    MongooseModule.forFeature([{ name: Conversation.name, schema: ConversationSchema }]),
  ],
  providers: [WebChatService],
  controllers: [WebChatController],
  exports: [WebChatService],
})
export class WebChatModule {}
