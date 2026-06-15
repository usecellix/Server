import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { FormulaModule } from '../formula/formula.module';
import { LlmModule } from '../llm/llm.module';
import { AgenticLoopService } from './agenticLoop.service';
import { OrchestratorService } from './orchestrator.service';
import { PlannerAgent } from './planner.agent';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import { CompletenessChecker } from './checkers/completeness.checker';
import { FormattingChecker } from './checkers/formatting.checker';
import { ToolBridgeService } from './tool-bridge.service';

@Module({
  imports: [AppConfigModule, LlmModule, FormulaModule],
  providers: [
    OrchestratorService,
    AgenticLoopService,
    PlannerAgent,
    ExecutorAgent,
    VerifierAgent,
    ToolBridgeService,
    CompletenessChecker,
    FormattingChecker,
  ],
  exports: [OrchestratorService, ToolBridgeService],
})
export class AgentsModule {}
