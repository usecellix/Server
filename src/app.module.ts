import { Module } from '@nestjs/common';
import { AuditModule } from './audit/audit.module';
import { AppConfigModule } from './config/app-config.module';
import { LoggingModule } from './common/logging/logging.module';
import { DatabaseModule } from './database/database.module';
import { DomainToolsModule } from './domain-tools/domain-tools.module';
import { ExcelAiModule } from './excel-ai/excel-ai.module';
import { HealthModule } from './health/health.module';
import { SheetsModule } from './sheets/sheets.module';

@Module({
  imports: [
    AppConfigModule,
    DatabaseModule,
    LoggingModule,
    AuditModule,
    HealthModule,
    ExcelAiModule,
    SheetsModule,
    DomainToolsModule,
  ],
})
export class AppModule {}
