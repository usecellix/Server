import { Injectable, Logger } from '@nestjs/common';
import { PlannerAgent } from './planner.agent';
import { AgenticLoopService } from './agenticLoop.service';
import { SseEmitter } from './sse.emitter';
import { Action, AgentRunOptions, PlannerOutput } from './types/agent.types';

@Injectable()
export class OrchestratorService {
  private readonly logger = new Logger(OrchestratorService.name);

  constructor(
    private readonly planner: PlannerAgent,
    private readonly agenticLoop: AgenticLoopService,
  ) {}

  /**
   * Plan mode: run only the PlannerAgent and return its structured plan without
   * executing any actions against the workbook.
   */
  async planOnly(opts: AgentRunOptions): Promise<PlannerOutput> {
    const { prompt, context, conversationHistory = [], promptContext, correlationId, routerAssumption } =
      opts;
    const resolvedCorrelationId = this.resolveCorrelationId(correlationId);
    return this.planner.plan(
      prompt,
      context,
      conversationHistory,
      promptContext,
      resolvedCorrelationId,
      routerAssumption,
    );
  }

  async run(opts: AgentRunOptions, emitter: SseEmitter): Promise<Action[]> {
    const {
      prompt,
      context,
      conversationHistory = [],
      promptContext,
      conversationId,
      correlationId,
      toolEmit,
      routerAssumption,
    } = opts;
    const resolvedCorrelationId = this.resolveCorrelationId(correlationId);

    emitter.send({ type: 'THINKING', message: 'Planning your request...' });
    const plan: PlannerOutput = await this.planner.plan(
      prompt,
      context,
      conversationHistory,
      promptContext,
      resolvedCorrelationId,
      routerAssumption,
    );

    if (plan.clarificationsNeeded.length > 0) {
      emitter.send({ type: 'CLARIFY', questions: plan.clarificationsNeeded });
      return [];
    }

    if (plan.confidence === 'low') {
      emitter.send({
        type: 'THINKING',
        message: `Low confidence — proceeding carefully (${plan.reasoning})`,
      });
    }

    emitter.send({
      type: 'CHECKPOINT',
      step: `${plan.subtasks.length} step${plan.subtasks.length > 1 ? 's' : ''} planned`,
    });

    const { actions: allActions, iterationsRun, verifierPassed } =
      await this.agenticLoop.run(prompt, plan.subtasks, context, emitter, {
        conversationId,
        correlationId: resolvedCorrelationId,
        toolEmit,
      });

    this.logger.log(
      `Agentic loop complete: ${allActions.length} actions, ${iterationsRun} iterations, verified: ${verifierPassed}`,
    );

    emitter.send({
      type: 'CHECKPOINT',
      step: `${allActions.length} actions ready for preview`,
    });

    return allActions;
  }

  private resolveCorrelationId(value?: string): string {
    const trimmed = value?.trim();
    if (!trimmed || trimmed === '-') return `req_${Date.now()}`;
    return trimmed;
  }
}
