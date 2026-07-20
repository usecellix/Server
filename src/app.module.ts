import { Module } from '@nestjs/common';
import { AuditModule } from './audit/audit.module';
import { AuthModule } from './auth/auth.module';
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
<<<<<<< HEAD
    LoggingModule,
=======
    AuthModule,
>>>>>>> 79b55a729d32439c8865d125c5c4c0c1a20e34a6
    AuditModule,
    HealthModule,
    ExcelAiModule,
    SheetsModule,
    DomainToolsModule,
  ],
})
export class AppModule {}
