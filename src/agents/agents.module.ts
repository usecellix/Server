import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { AppConfigModule } from '../config/app-config.module';
import { LoggingModule } from '../common/logging/logging.module';
import { FormulaModule } from '../formula/formula.module';
import { LlmModule } from '../llm/llm.module';
import { AgenticLoopService } from './agenticLoop.service';
import { AgentRunStateService } from './agent-run-state.service';
import { AgentRun, AgentRunSchema } from './schemas/agent-run.schema';
import { OrchestratorService } from './orchestrator.service';
import { PlannerAgent } from './planner.agent';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import { CompletenessChecker } from './checkers/completeness.checker';
import { FormattingChecker } from './checkers/formatting.checker';
import { SemanticFormulaChecker } from './checkers/semantic-formula.checker';
import { OverwriteOccupancyChecker } from './checkers/overwrite-occupancy.checker';
import { StructuralIntentChecker } from './checkers/structural-intent.checker';
import { ComputedColumnChecker } from './checkers/computed-column.checker';
import { SpecConformanceChecker } from './checkers/spec-conformance.checker';
import { SpecExtractorAgent } from './spec-extractor.agent';
import { ToolBridgeService } from './tool-bridge.service';
import { StructuredLogger } from './logging/structured-logger';

@Module({
  imports: [
    AppConfigModule,
    LlmModule,
    FormulaModule,
    LoggingModule,
    MongooseModule.forFeature([{ name: AgentRun.name, schema: AgentRunSchema }]),
  ],
  providers: [
    OrchestratorService,
    AgenticLoopService,
    AgentRunStateService,
    PlannerAgent,
    ExecutorAgent,
    VerifierAgent,
    ToolBridgeService,
    StructuredLogger,
    CompletenessChecker,
    FormattingChecker,
    SemanticFormulaChecker,
    OverwriteOccupancyChecker,
    StructuralIntentChecker,
    ComputedColumnChecker,
    SpecConformanceChecker,
    SpecExtractorAgent,
  ],
  exports: [
    OrchestratorService,
    AgentRunStateService,
    ToolBridgeService,
    ExecutorAgent,
    VerifierAgent,
    StructuredLogger,
  ],
})
export class AgentsModule {}
