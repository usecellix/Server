import { Module } from '@nestjs/common';
import { AuditModule } from './audit/audit.module';
import { AuthModule } from './auth/auth.module';
import { AppConfigModule } from './config/app-config.module';
import { CreditModule } from './credit/credit.module';
import { LoggingModule } from './common/logging/logging.module';
import { DatabaseModule } from './database/database.module';
import { DomainToolsModule } from './domain-tools/domain-tools.module';
import { ExcelAiModule } from './excel-ai/excel-ai.module';
import { GstReconModule } from './gst-recon/gst-recon.module';
import { HealthModule } from './health/health.module';
import { SheetsModule } from './sheets/sheets.module';

@Module({
  imports: [
    AppConfigModule,
    DatabaseModule,
    LoggingModule,
    AuthModule,
    AuditModule,
    CreditModule,
    HealthModule,
    ExcelAiModule,
    SheetsModule,
    DomainToolsModule,
    GstReconModule,
  ],
})
export class AppModule {}
