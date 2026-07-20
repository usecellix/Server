import { Module } from '@nestjs/common';
import { DOMAIN_TOOL_REGISTRY, domainToolRegistry } from './registry';

/**
 * Scaffolding module for deterministic domain tools (GST / TDS / recon / accounting).
 * Stub implementations throw until CA-reviewed logic is signed off — do not expose
 * stub outputs to end users as compliance-correct results.
 */
@Module({
  providers: [
    {
      provide: DOMAIN_TOOL_REGISTRY,
      useValue: domainToolRegistry,
    },
  ],
  exports: [DOMAIN_TOOL_REGISTRY],
})
export class DomainToolsModule {}
