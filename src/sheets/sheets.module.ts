import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { ModelRouter } from '../excel-ai/llm/model-router';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import { MultiSheetService } from './multi-sheet.service';
import { SheetsController } from './sheets.controller';

@Module({
  imports: [AppConfigModule],
  providers: [ModelRouter, MultiSheetService, OpenRouterService],
  controllers: [SheetsController],
  exports: [MultiSheetService],
})
export class SheetsModule {}
