import { Module } from '@nestjs/common';
import { DomainToolsModule } from '../domain-tools/domain-tools.module';
import { AuditModule } from '../audit/audit.module';
import { GstReconController } from './gst-recon.controller';
import { GstReconService } from './gst-recon.service';
import { GstReconOrchestrator } from './orchestrator';

@Module({
  imports: [DomainToolsModule, AuditModule],
  controllers: [GstReconController],
  providers: [GstReconService, GstReconOrchestrator],
  exports: [GstReconService, GstReconOrchestrator],
})
export class GstReconModule {}
