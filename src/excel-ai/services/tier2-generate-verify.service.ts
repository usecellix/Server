import { Injectable, Logger } from '@nestjs/common';
import { ExecutorAgent } from '../../agents/executor.agent';
import { VerifierAgent } from '../../agents/verifier.agent';
import { ToolBridgeService } from '../../agents/tool-bridge.service';
import {
  Action,
  ExecutorOutput,
  SubTask,
  VerifierOutput,
  WorkbookContext,
} from '../../agents/types/agent.types';
import { mergeRangeIntoSheet } from '../../agents/utils/range-merge.util';
import { FormulaValidatorService } from '../../formula/formula-validator.service';
import { SourceRef } from '../../domain-tools/types/domain-tool.types';
import { buildWorkbookSourceRefsFromActions } from '../../audit/utils/provenance.util';
import { SheetAction } from '../types/sheet-actions.types';
import { assertTier2VerifierMandatory } from '../utils/tier2-verifier.guard';
import {
  applyPriorSourceRangeToChartActions,
  ensureRequestedChartColorScheme,
  extractSuggestedSourceRange,
  patchChartSourceRanges,
} from '../utils/chart-range-correction.util';
import { TurnActionRecord } from '../utils/turn-action-history.util';

export interface Tier2ToolBridgeOptions {
  conversationId: string;
  toolEmit: (event: string, data: Record<string, unknown>) => void;
}

export interface Tier2Result {
  actions: SheetAction[];
  answer: string;
  verifierPassed: boolean;
  verifierSkipped: false;
  failureReason?: string;
  durationMs: number;
  /** Workbook citations for formula precedents (Tier 2 writes) */
  sourceRefs: SourceRef[];
  /** True when the single bounded verifier-feedback retry ran. */
  retried?: boolean;
  /**
   * Spec 18 Bug 4 — true when retry returned toolRequest and we made one
   * additional Executor call with the fetched range data.
   */
  toolFollowUp?: boolean;
  /** True when sourceRange/color were patched deterministically (no extra LLM). */
  deterministicPatch?: boolean;
}

export interface Tier2GenerateOnlyResult {
  actions: SheetAction[];
  answer: string;
  blockedReason?: string;
  durationMs: number;
}

/** Spec 18: at most retry + one tool-informed follow-up. */
const MAX_CORRECTION_EXECUTOR_CALLS = 2;

@Injectable()
export class Tier2GenerateVerifyService {
  private readonly logger = new Logger(Tier2GenerateVerifyService.name);

  constructor(
    private readonly executorAgent: ExecutorAgent,
    private readonly verifierAgent: VerifierAgent,
    private readonly formulaValidator: FormulaValidatorService,
    private readonly toolBridge: ToolBridgeService,
  ) {}

  /** Plan mode: generate a proposal via ExecutorAgent only — no Verifier, no ChangeSet. */
  async generateOnly(
    message: string,
    actionHint: string,
    workbookContext: WorkbookContext,
    correlationId = `req_${Date.now()}`,
  ): Promise<Tier2GenerateOnlyResult> {
    const startedAt = Date.now();
    const subtask = this.buildSyntheticSubtask(message, actionHint, workbookContext);

    const executorResult = await this.executorAgent.execute(
      subtask,
      workbookContext,
      [],
      correlationId,
    );

    const hardcodeCheck = this.formulaValidator.checkNoHardcodedLiterals(executorResult.actions);
    if (!hardcodeCheck.passed) {
      this.logger.warn(
        `Tier 2 generateOnly blocked by hardcode lint: ${hardcodeCheck.reason ?? 'unknown'}`,
      );
      return {
        actions: executorResult.actions,
        answer: `Proposal blocked: ${hardcodeCheck.reason ?? 'Hardcoded literal detected'}`,
        blockedReason: hardcodeCheck.reason ?? 'Hardcoded literal detected',
        durationMs: Date.now() - startedAt,
      };
    }

    const durationMs = Date.now() - startedAt;
    this.logger.log(
      `Tier 2 generateOnly tier=2 actionHint=${actionHint} actions=${executorResult.actions.length} durationMs=${durationMs}`,
    );

    return {
      actions: executorResult.actions,
      answer: this.buildProposalAnswer(executorResult.actions, actionHint),
      durationMs,
    };
  }

  async execute(
    message: string,
    actionHint: string,
    workbookContext: WorkbookContext,
    correlationId = `req_${Date.now()}`,
    toolOptions?: Tier2ToolBridgeOptions,
  ): Promise<Tier2Result> {
    const startedAt = Date.now();
    const subtask = this.buildSyntheticSubtask(message, actionHint, workbookContext);

    let executorResult = await this.executorAgent.execute(
      subtask,
      workbookContext,
      [],
      correlationId,
    );

    let deterministicPatch = false;
    const afterFirst = this.applyDeterministicChartFixes(
      message,
      executorResult,
      workbookContext,
    );
    executorResult = afterFirst.result;
    deterministicPatch = deterministicPatch || afterFirst.applied;

    const lintOrPreApplyFail = this.runDeterministicGates(executorResult, workbookContext, startedAt);
    if (lintOrPreApplyFail) return lintOrPreApplyFail;

    assertTier2VerifierMandatory({ usedShouldSkipVerifier: false });

    let workingContext = workbookContext;
    let preApply = this.formulaValidator.validatePreApply(
      executorResult.actions,
      workingContext,
    );
    let verifierResult = await this.verifierAgent.verify(
      message,
      [subtask],
      { [subtask.id]: executorResult.actions },
      workingContext,
      this.formulaValidator.summarizeForVerifier([preApply]),
      correlationId,
    );

    let retried = false;
    let toolFollowUp = false;
    let correctionExecutorCalls = 0;

    if (!verifierResult.passed) {
      // Spec 18 Bug 1: exactly one retry with verifier feedback — not a blind re-roll.
      retried = true;
      correctionExecutorCalls += 1;
      const feedback = this.formatVerifierFeedbackForRetry(verifierResult);
      this.logger.warn(`Tier 2 verifier failed — retrying once with feedback: ${feedback.slice(0, 200)}`);

      let correctionContext: WorkbookContext = {
        ...workingContext,
        verifierFeedback: feedback,
        verifierIssues: verifierResult.issues,
      };

      executorResult = await this.executorAgent.retryStep(
        {
          originalStep: subtask,
          attempt: 1,
          maxAttempts: 1,
          verifierFeedback: feedback,
        },
        correctionContext,
        [],
        correlationId,
      );

      // Spec 18 Bug 4: toolRequest during retry is data-gathering, not the correction itself.
      // Cap correction-related Executor calls at 2 (retry + one tool-informed follow-up).
      if (
        this.needsToolInformedFollowUp(executorResult) &&
        correctionExecutorCalls < MAX_CORRECTION_EXECUTOR_CALLS &&
        toolOptions
      ) {
        const followUp = await this.resolveToolAndFollowUp(
          subtask,
          correctionContext,
          executorResult,
          toolOptions,
          correlationId,
        );
        if (followUp) {
          executorResult = followUp.result;
          correctionContext = followUp.context;
          workingContext = followUp.context;
          toolFollowUp = true;
          correctionExecutorCalls += 1;
        }
      } else if (this.needsToolInformedFollowUp(executorResult) && !toolOptions) {
        this.logger.warn(
          'Tier 2 retry returned toolRequest but no tool bridge options were provided — cannot follow up',
        );
      }

      const afterRetry = this.applyDeterministicChartFixes(
        message,
        executorResult,
        workingContext,
      );
      executorResult = afterRetry.result;
      deterministicPatch = deterministicPatch || afterRetry.applied;

      // If the verifier already suggested a concrete range, apply it before re-verify
      // so a color-only first failure doesn't leave a known-bad A4 range untouched.
      const suggestedFromFirst = extractSuggestedSourceRange(verifierResult);
      if (suggestedFromFirst) {
        const patched = patchChartSourceRanges(executorResult.actions, suggestedFromFirst);
        if (patched.patched) {
          this.logger.log(
            `Tier 2: applying verifier-suggested sourceRange ${suggestedFromFirst} before re-verify`,
          );
          executorResult = { ...executorResult, actions: patched.actions };
          deterministicPatch = true;
        }
      }

      const retryGate = this.runDeterministicGates(executorResult, workingContext, startedAt);
      if (retryGate) {
        return { ...retryGate, retried: true, toolFollowUp, deterministicPatch };
      }

      preApply = this.formulaValidator.validatePreApply(
        executorResult.actions,
        workingContext,
      );
      assertTier2VerifierMandatory({ usedShouldSkipVerifier: false });
      verifierResult = await this.verifierAgent.verify(
        message,
        [subtask],
        { [subtask.id]: executorResult.actions },
        workingContext,
        this.formulaValidator.summarizeForVerifier([preApply]),
        correlationId,
      );
    }

    // After the bounded LLM correction path, still apply a concrete range suggestion
    // from this verify pass (no extra Executor call — mechanical patch + one re-verify).
    if (!verifierResult.passed) {
      const suggested = extractSuggestedSourceRange(verifierResult);
      if (suggested) {
        const patched = patchChartSourceRanges(executorResult.actions, suggested);
        if (patched.patched) {
          this.logger.log(
            `Tier 2: deterministic sourceRange patch ${suggested} after verifier failure`,
          );
          executorResult = { ...executorResult, actions: patched.actions };
          deterministicPatch = true;

          const gate = this.runDeterministicGates(executorResult, workingContext, startedAt);
          if (!gate) {
            preApply = this.formulaValidator.validatePreApply(
              executorResult.actions,
              workingContext,
            );
            assertTier2VerifierMandatory({ usedShouldSkipVerifier: false });
            verifierResult = await this.verifierAgent.verify(
              message,
              [subtask],
              { [subtask.id]: executorResult.actions },
              workingContext,
              this.formulaValidator.summarizeForVerifier([preApply]),
              correlationId,
            );
          }
        }
      }
    }

    const durationMs = Date.now() - startedAt;
    this.logger.log(
      `Tier 2 complete tier=2 actionHint=${actionHint} verifierPassed=${verifierResult.passed} retried=${retried} toolFollowUp=${toolFollowUp} deterministicPatch=${deterministicPatch} correctionCalls=${correctionExecutorCalls} durationMs=${durationMs}`,
    );

    if (!verifierResult.passed) {
      const failureMessage = this.buildVerifierFailureAnswer(verifierResult);
      return {
        actions: executorResult.actions,
        answer: failureMessage,
        verifierPassed: false,
        verifierSkipped: false,
        failureReason: failureMessage,
        durationMs,
        sourceRefs: this.buildSourceRefs(executorResult.actions, workingContext),
        retried,
        toolFollowUp,
        deterministicPatch,
      };
    }

    return {
      actions: executorResult.actions,
      answer: this.buildSuccessAnswer(executorResult.actions, actionHint),
      verifierPassed: true,
      verifierSkipped: false,
      durationMs,
      sourceRefs: this.buildSourceRefs(executorResult.actions, workingContext),
      retried,
      toolFollowUp,
      deterministicPatch,
    };
  }

  private applyDeterministicChartFixes(
    message: string,
    result: ExecutorOutput,
    context: WorkbookContext,
  ): { result: ExecutorOutput; applied: boolean } {
    let actions = result.actions ?? [];
    let applied = false;

    const color = ensureRequestedChartColorScheme(message, actions);
    actions = color.actions;
    applied = applied || color.applied;

    const prior = applyPriorSourceRangeToChartActions(
      actions,
      context.priorTurnActions as TurnActionRecord[] | undefined,
      message,
    );
    actions = prior.actions;
    if (prior.applied) {
      this.logger.log(
        `Tier 2: applying prior-turn sourceRange ${prior.sourceRange} for follow-up chart request`,
      );
      applied = true;
    }

    if (!applied) return { result, applied: false };
    return { result: { ...result, actions }, applied: true };
  }

  private needsToolInformedFollowUp(result: ExecutorOutput): boolean {
    return Boolean(result.toolRequest) && (!result.actions || result.actions.length === 0);
  }

  private async resolveToolAndFollowUp(
    subtask: SubTask,
    context: WorkbookContext,
    toolResult: ExecutorOutput,
    toolOptions: Tier2ToolBridgeOptions,
    correlationId: string,
  ): Promise<{ result: ExecutorOutput; context: WorkbookContext } | null> {
    const toolRequest = toolResult.toolRequest;
    if (!toolRequest) return null;

    this.logger.log(
      `Tier 2 Bug 4: resolving toolRequest get_range_data(${toolRequest.sheet}, ${toolRequest.range}) then one follow-up Executor call`,
    );

    try {
      const fetched = await this.toolBridge.waitForRangeData(
        toolOptions.conversationId,
        toolRequest,
        toolOptions.toolEmit,
      );

      if (fetched.error || !fetched.values?.length) {
        this.logger.warn(
          `Tier 2 tool follow-up aborted: ${fetched.error ?? 'empty values'}`,
        );
        return null;
      }

      const mergedContext = this.mergeFetchedRange(context, toolRequest, fetched.values);
      const result = await this.executorAgent.execute(
        subtask,
        mergedContext,
        [],
        correlationId,
      );
      return { result, context: mergedContext };
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      this.logger.warn(`Tier 2 tool follow-up failed: ${reason}`);
      return null;
    }
  }

  private mergeFetchedRange(
    context: WorkbookContext,
    toolRequest: NonNullable<ExecutorOutput['toolRequest']>,
    values: unknown[][],
  ): WorkbookContext {
    const sheets = context.sheets.map((sheet) => {
      if (sheet.name !== toolRequest.sheet) return sheet;
      return mergeRangeIntoSheet(sheet, toolRequest.range, values);
    });

    return {
      ...context,
      sheets,
      fetchedRanges: [
        ...(context.fetchedRanges ?? []),
        {
          sheet: toolRequest.sheet,
          range: toolRequest.range,
          rowCount: values.length,
        },
      ],
    };
  }

  private runDeterministicGates(
    executorResult: ExecutorOutput,
    workbookContext: WorkbookContext,
    startedAt: number,
  ): Tier2Result | null {
    const hardcodeCheck = this.formulaValidator.checkNoHardcodedLiterals(executorResult.actions);
    if (!hardcodeCheck.passed) {
      this.logger.warn(
        `Tier 2 hardcode lint blocked verifier call: ${hardcodeCheck.reason ?? 'unknown'}`,
      );
      return this.buildFailureResult(
        executorResult,
        hardcodeCheck.reason ?? 'Hardcoded literal detected',
        startedAt,
        workbookContext,
      );
    }

    const preApply = this.formulaValidator.validatePreApply(
      executorResult.actions,
      workbookContext,
    );
    if (!preApply.passed) {
      const reason = this.formulaValidator.formatFeedback(preApply.issues);
      this.logger.warn(`Tier 2 pre-apply formula validation failed: ${reason}`);
      return this.buildFailureResult(executorResult, reason, startedAt, workbookContext);
    }

    return null;
  }

  private formatVerifierFeedbackForRetry(verifierResult: VerifierOutput): string {
    const fromIssues = verifierResult.issues
      .map((issue) => {
        const suggestion = issue.suggestion?.trim();
        return suggestion
          ? `${issue.description} Suggestion: ${suggestion}`
          : issue.description;
      })
      .filter(Boolean);

    if (fromIssues.length > 0) {
      return fromIssues.join('\n');
    }

    const fromSubtasks = verifierResult.subtaskResults
      .filter((r) => !r.passed)
      .flatMap((r) =>
        r.issues.length > 0
          ? r.issues.map((issue) =>
              issue.suggestion
                ? `${issue.description} Suggestion: ${issue.suggestion}`
                : issue.description,
            )
          : [r.feedback],
      )
      .filter(Boolean);

    if (fromSubtasks.length > 0) {
      return fromSubtasks.join('\n');
    }

    return verifierResult.feedback || 'Previous attempt failed verification — correct the actions.';
  }

  private buildVerifierFailureAnswer(verifierResult: VerifierOutput): string {
    const suggestions = this.collectSuggestions(verifierResult);
    const base = verifierResult.feedback?.trim() || 'Verification failed';
    if (suggestions.length === 0) {
      return `Verification failed: ${base}`;
    }
    return `Verification failed: ${base} Try: ${suggestions.join('; ')}.`;
  }

  private collectSuggestions(verifierResult: VerifierOutput): string[] {
    const fromIssues = verifierResult.issues
      .map((i) => i.suggestion?.trim())
      .filter((s): s is string => Boolean(s));
    const fromSubtasks = verifierResult.subtaskResults.flatMap((r) =>
      r.issues.map((i) => i.suggestion?.trim()).filter((s): s is string => Boolean(s)),
    );
    return [...new Set([...fromIssues, ...fromSubtasks])];
  }

  private buildSourceRefs(actions: Action[], workbookContext: WorkbookContext): SourceRef[] {
    const workbookId = workbookContext.activeSheetName || 'workbook';
    return buildWorkbookSourceRefsFromActions(
      actions,
      workbookId,
      workbookContext.activeSheetName,
    );
  }

  private buildSyntheticSubtask(
    message: string,
    actionHint: string,
    workbookContext: WorkbookContext,
  ): SubTask {
    return {
      id: 's1',
      description: `[${actionHint}] ${message}`,
      targetSheet: workbookContext.activeSheetName,
      dependsOn: [],
      estimatedActions: 1,
      ...(actionHint === 'CREATE_CHART' || actionHint === 'CREATE_CHARTS'
        ? { suggestedActionType: 'CREATE_CHART' }
        : {}),
    };
  }

  private buildFailureResult(
    executorResult: ExecutorOutput,
    reason: string,
    startedAt: number,
    workbookContext?: WorkbookContext,
  ): Tier2Result {
    return {
      actions: executorResult.actions,
      answer: `Verification failed before apply: ${reason}`,
      verifierPassed: false,
      verifierSkipped: false,
      failureReason: reason,
      durationMs: Date.now() - startedAt,
      sourceRefs: workbookContext
        ? this.buildSourceRefs(executorResult.actions, workbookContext)
        : [],
    };
  }

  private buildSuccessAnswer(actions: Action[], actionHint: string): string {
    const count = actions.length;
    if (count === 0) {
      return `Tier 2 (${actionHint}) completed with no actions.`;
    }
    if (count === 1) {
      return `Prepared **1** verified change for your sheet (${actionHint}).`;
    }
    return `Prepared **${count}** verified changes for your sheet (${actionHint}).`;
  }

  private buildProposalAnswer(actions: Action[], actionHint: string): string {
    const count = actions.length;
    if (count === 0) {
      return `No proposed changes for ${actionHint}.`;
    }
    if (count === 1) {
      return `Proposed **1** change (${actionHint}) — review the plan, then run as Action to verify and apply.`;
    }
    return `Proposed **${count}** changes (${actionHint}) — review the plan, then run as Action to verify and apply.`;
  }
}
