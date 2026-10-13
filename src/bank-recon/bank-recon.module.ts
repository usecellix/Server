import { Module } from '@nestjs/common';
import { BankReconController } from './bank-recon.controller';
import { BankReconService } from './bank-recon.service';

@Module({
  controllers: [BankReconController],
  providers: [BankReconService],
  exports: [BankReconService],
})
export class BankReconModule {}
