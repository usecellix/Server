import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { BankStatementController } from './bank-statement.controller';
import { BankStatementService } from './bank-statement.service';

/**
 * Bank statement import. Depends on nothing but config (for the auth guard):
 * the parsing lives in `domain-tools/ingestion` as plain functions, and no LLM
 * module is imported because this path never calls one.
 */
@Module({
  imports: [AppConfigModule],
  controllers: [BankStatementController],
  providers: [BankStatementService],
  exports: [BankStatementService],
})
export class BankStatementModule {}
