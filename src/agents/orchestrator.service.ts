import { Injectable, Logger } from '@nestjs/common';
import { PlannerAgent } from './planner.agent';
import { AgenticLoopService } from './agenticLoop.service';
import { SseEmitter } from './sse.emitter';
import { Action, AgentRunOptions, PlannerOutput } from './types/agent.types';
import { annotateExplicitOverwriteConfirmation } from '../excel-ai/utils/overwrite-confirmation.util';
import { annotateClearIntentOverwrite } from './utils/clear-intent-overwrite.util';
import { pruneSpuriousAddSheetActions } from './utils/compound-action.util';

export interface OrchestratorRunResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
  clarificationRequested: boolean;
  completedSubtasks: Array<{ subtaskId: string; actions: Action[]; verified: boolean }>;
  failedSubtask: { subtaskId: string; reason: string } | null;
  partialProgress: boolean;
  /**
   * The plan's own natural-language statements of intent, surfaced so the
   * Accept card can describe what the build DOES rather than enumerate the
   * mechanical actions it emits. The Planner already writes exactly the right
   * text ("Create sheet 'January' ... and set A1:J1 headers [...]"); it was
   * previously streamed as a transient `status` event and then discarded,
   * leaving the card to render ~40 action bullets instead. TASKS.md #149.
   */
  planSubtasks: Array<{ id: string; description: string; targetSheet: string }>;
  /**
   * Subtasks the Planner produced that the Executor delivered no actions for.
   * Non-empty means the build is incomplete relative to its own plan — the
   * `estimatedActions` circularity `CODEBASE_ANALYSIS.md` §3.15 describes,
   * caught here at the one place both numbers are known. TASKS.md #155.
   */
  undeliveredSubtasks: Array<{ id: string; description: string; targetSheet: string }>;
}

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
    const {
      prompt,
      context,
      conversationHistory = [],
      promptContext,
      correlationId,
      routerAssumption,
      complexity,
    } = opts;
    const resolvedCorrelationId = this.resolveCorrelationId(correlationId);
    return this.planner.plan(
      prompt,
      context,
      conversationHistory,
      promptContext,
      resolvedCorrelationId,
      routerAssumption,
      complexity,
    );
  }

  async run(opts: AgentRunOptions, emitter: SseEmitter): Promise<Action[]> {
    const result = await this.runDetailed(opts, emitter);
    return result.actions;
  }

  async runDetailed(
    opts: AgentRunOptions,
    emitter: SseEmitter,
  ): Promise<OrchestratorRunResult> {
    const {
      prompt,
      context,
      conversationHistory = [],
      promptContext,
      conversationId,
      correlationId,
      toolEmit,
      routerAssumption,
      complexity,
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
      complexity,
    );

    if (plan.clarificationsNeeded.length > 0) {
      emitter.send({ type: 'CLARIFY', questions: plan.clarificationsNeeded });
      return {
        actions: [],
        iterationsRun: 0,
        verifierPassed: false,
        clarificationRequested: true,
        completedSubtasks: [],
        failedSubtask: null,
        partialProgress: false,
        planSubtasks: [],
        undeliveredSubtasks: [],
      };
    }

    if (plan.confidence === 'low') {
      const questions =
        plan.clarificationsNeeded.length > 0
          ? plan.clarificationsNeeded
          : [
              plan.reasoning?.trim() ||
                'This request is ambiguous — what exactly should I change in the workbook?',
            ];
      emitter.send({ type: 'CLARIFY', questions });
      return {
        actions: [],
        iterationsRun: 0,
        verifierPassed: false,
        clarificationRequested: true,
        completedSubtasks: [],
        failedSubtask: null,
        partialProgress: false,
        planSubtasks: [],
        undeliveredSubtasks: [],
      };
    }

    emitter.send({
      type: 'CHECKPOINT',
      step: `${plan.subtasks.length} step${plan.subtasks.length > 1 ? 's' : ''} planned`,
    });

    const {
      actions: allActions,
      iterationsRun,
      verifierPassed,
      completedSubtasks,
      failedSubtask,
      partialProgress,
    } = await this.agenticLoop.run(prompt, plan.subtasks, context, emitter, {
      conversationId,
      correlationId: resolvedCorrelationId,
      toolEmit,
    });

    this.logger.log(
      `Agentic loop complete: ${allActions.length} actions, ${iterationsRun} iterations, verified: ${verifierPassed}, partial: ${partialProgress}`,
    );

    emitter.send({
      type: 'CHECKPOINT',
      step: `${allActions.length} actions ready for preview`,
    });

    const pruned = pruneSpuriousAddSheetActions(allActions);
    const clearAnnotated = annotateClearIntentOverwrite(pruned, prompt);
    const overwriteAnnotated = annotateExplicitOverwriteConfirmation(
      clearAnnotated,
      prompt,
      context.priorTurnActions ?? [],
    );

    return {
      actions: overwriteAnnotated,
      iterationsRun,
      verifierPassed,
      clarificationRequested: false,
      completedSubtasks,
      failedSubtask,
      partialProgress,
      // ONLY subtasks that actually produced actions. The Accept card renders
      // these as its promises (TASKS.md #149), so a subtask the Executor
      // silently emitted nothing for must not appear there — a live smoke test
      // caught the card promising "Write Consolidated Transactions header at
      // Main!A18" for a run whose actions never touched row 18. Describing the
      // PLAN rather than the DELIVERY is precisely the false-completeness shape
      // CODEBASE_ANALYSIS.md §3.7 keeps re-teaching. TASKS.md #155.
      planSubtasks: plan.subtasks
        .filter((s) =>
          completedSubtasks.some((c) => c.subtaskId === s.id && c.actions.length > 0),
        )
        .map((s) => ({
          id: s.id,
          description: s.description,
          targetSheet: s.targetSheet,
        })),
      /** Planned but delivered nothing — surfaced so the gap can be reported. */
      undeliveredSubtasks: plan.subtasks
        .filter(
          (s) =>
            !completedSubtasks.some((c) => c.subtaskId === s.id && c.actions.length > 0),
        )
        .map((s) => ({
          id: s.id,
          description: s.description,
          targetSheet: s.targetSheet,
        })),
    };
  }

  private resolveCorrelationId(value?: string): string {
    const trimmed = value?.trim();
    if (!trimmed || trimmed === '-') return `req_${Date.now()}`;
    return trimmed;
  }
}
