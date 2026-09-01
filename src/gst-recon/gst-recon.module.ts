import { Module } from '@nestjs/common';
import { DomainToolsModule } from '../domain-tools/domain-tools.module';
import { GstReconController } from './gst-recon.controller';
import { GstReconService } from './gst-recon.service';

@Module({
  imports: [DomainToolsModule],
  controllers: [GstReconController],
  providers: [GstReconService],
  exports: [GstReconService],
})
export class GstReconModule {}
