import { Injectable, Logger } from '@nestjs/common';
import { ExecutorAgent } from '../../agents/executor.agent';
import { VerifierAgent } from '../../agents/verifier.agent';
import { Action, ExecutorOutput, SubTask, WorkbookContext } from '../../agents/types/agent.types';
import { FormulaValidatorService } from '../../formula/formula-validator.service';
import { SourceRef } from '../../domain-tools/types/domain-tool.types';
import { buildWorkbookSourceRefsFromActions } from '../../audit/utils/provenance.util';
import { SheetAction } from '../types/sheet-actions.types';
import { assertTier2VerifierMandatory } from '../utils/tier2-verifier.guard';

export interface Tier2Result {
  actions: SheetAction[];
  answer: string;
  verifierPassed: boolean;
  verifierSkipped: false;
  failureReason?: string;
  durationMs: number;
  /** Workbook citations for formula precedents (Tier 2 writes) */
  sourceRefs: SourceRef[];
}

export interface Tier2GenerateOnlyResult {
  actions: SheetAction[];
  answer: string;
  blockedReason?: string;
  durationMs: number;
}

@Injectable()
export class Tier2GenerateVerifyService {
  private readonly logger = new Logger(Tier2GenerateVerifyService.name);

  constructor(
    private readonly executorAgent: ExecutorAgent,
    private readonly verifierAgent: VerifierAgent,
    private readonly formulaValidator: FormulaValidatorService,
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
  ): Promise<Tier2Result> {
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

    assertTier2VerifierMandatory({ usedShouldSkipVerifier: false });

    const verifierResult = await this.verifierAgent.verify(
      message,
      [subtask],
      { [subtask.id]: executorResult.actions },
      workbookContext,
      this.formulaValidator.summarizeForVerifier([preApply]),
      correlationId,
    );

    const durationMs = Date.now() - startedAt;
    this.logger.log(
      `Tier 2 complete tier=2 actionHint=${actionHint} verifierPassed=${verifierResult.passed} durationMs=${durationMs}`,
    );

    return {
      actions: executorResult.actions,
      answer: verifierResult.passed
        ? this.buildSuccessAnswer(executorResult.actions, actionHint)
        : `Verification failed: ${verifierResult.feedback}`,
      verifierPassed: verifierResult.passed,
      verifierSkipped: false,
      failureReason: verifierResult.passed ? undefined : verifierResult.feedback,
      durationMs,
      sourceRefs: this.buildSourceRefs(executorResult.actions, workbookContext),
    };
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
