import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { BankReconController } from './bank-recon.controller';
import { BankReconService } from './bank-recon.service';

@Module({
  imports: [AppConfigModule],
  controllers: [BankReconController],
  providers: [BankReconService],
  exports: [BankReconService],
})
export class BankReconModule {}
