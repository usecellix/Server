import { Module } from '@nestjs/common';
import { AppConfigModule } from '../config/app-config.module';
import { LoggingModule } from '../common/logging/logging.module';
import { FormulaModule } from '../formula/formula.module';
import { LlmModule } from '../llm/llm.module';
import { AgenticLoopService } from './agenticLoop.service';
import { OrchestratorService } from './orchestrator.service';
import { PlannerAgent } from './planner.agent';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import { CompletenessChecker } from './checkers/completeness.checker';
import { FormattingChecker } from './checkers/formatting.checker';
import { SemanticFormulaChecker } from './checkers/semantic-formula.checker';
import { ToolBridgeService } from './tool-bridge.service';
import { StructuredLogger } from './logging/structured-logger';

@Module({
  imports: [AppConfigModule, LlmModule, FormulaModule, LoggingModule],
  providers: [
    OrchestratorService,
    AgenticLoopService,
    PlannerAgent,
    ExecutorAgent,
    VerifierAgent,
    ToolBridgeService,
    StructuredLogger,
    CompletenessChecker,
    FormattingChecker,
    SemanticFormulaChecker,
  ],
  exports: [OrchestratorService, ToolBridgeService, ExecutorAgent, VerifierAgent, StructuredLogger],
})
export class AgentsModule {}
