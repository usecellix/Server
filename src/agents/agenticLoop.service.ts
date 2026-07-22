import { Injectable, Logger } from '@nestjs/common';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import {
  Action,
  ExecutorOutput,
  SubTask,
  VerifierIssue,
  VerifierOutput,
  WorkbookContext,
} from './types/agent.types';
import { SseEmitter } from './sse.emitter';
import { FormulaAnalyzer } from '../formula/formula.analyzer';
import { FormulaValidatorService } from '../formula/formula-validator.service';
import { FormulaValidationResult } from '../formula/formula.types';
import { buildShadowWorkbook, shadowToWorkbookContext } from '../virtual/shadowWorkbook';
import { ShadowWorkbook } from '../virtual/shadowWorkbook.types';
import { virtualApply } from '../virtual/virtualApply';
import { ToolBridgeService } from './tool-bridge.service';
import { mergeRangeIntoSheet } from './utils/range-merge.util';
import { computeExecutionWaves } from './utils/task-graph.util';
import { CompletenessChecker } from './checkers/completeness.checker';
import { FormattingChecker } from './checkers/formatting.checker';
import { SemanticFormulaChecker } from './checkers/semantic-formula.checker';
import { CheckerResult, mergeCheckerResults } from './checkers/checker.types';
import { buildDeterministicSubtaskActions } from './utils/compound-action.util';
import { StepRetryExhaustedError } from './errors';
import { StepRetryContext } from './types/verifier.types';
import { StructuredLogger } from './logging/structured-logger';
import { shouldSkipVerifier } from './verifier-skip.policy';
import { isExecutorBlockedSignal } from './utils/verifier-partial-parse.util';

export interface AgenticLoopOptions {
  conversationId?: string;
  correlationId?: string;
  toolEmit?: (event: string, data: Record<string, unknown>) => void;
  parseFailureTracker?: { hadFailure: boolean };
}

export interface CompletedSubtaskResult {
  subtaskId: string;
  actions: Action[];
  verified: boolean;
}

export interface FailedSubtaskResult {
  subtaskId: string;
  reason: string;
}

export interface AgenticLoopResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
  completedSubtasks: CompletedSubtaskResult[];
  failedSubtask: FailedSubtaskResult | null;
  /** True when some subtasks completed but the full chain did not verify/pass. */
  partialProgress: boolean;
}

interface SubtaskActionState {
  subtask: SubTask;
  actions: Action[];
  completed: boolean;
  verified?: boolean;
  failedReason?: string;
}

type RetryContext = Pick<
  WorkbookContext,
  'verifierFeedback' | 'verifierIssues' | 'formulaValidationFeedback' | 'formulaValidationIssues'
>;

@Injectable()
export class AgenticLoopService {
  private readonly logger = new Logger(AgenticLoopService.name);
  private readonly MAX_ITERATIONS_PER_SUBTASK = 10;
  private readonly MAX_STEP_RETRIES = 2;
  private readonly MAX_FORMULA_RETRIES = 2;
  private readonly MAX_TOOL_REQUESTS = 5;
  private readonly TIMEOUT_MS = 300_000;

  constructor(
    private readonly executor: ExecutorAgent,
    private readonly verifier: VerifierAgent,
    private readonly formulaAnalyzer: FormulaAnalyzer,
    private readonly formulaValidator: FormulaValidatorService,
    private readonly toolBridge: ToolBridgeService,
    private readonly completenessChecker: CompletenessChecker,
    private readonly formattingChecker: FormattingChecker,
    private readonly semanticFormulaChecker: SemanticFormulaChecker = new SemanticFormulaChecker(),
    private readonly structuredLogger: StructuredLogger = new StructuredLogger(),
  ) {}

  async run(
    originalPrompt: string,
    subtasks: SubTask[],
    context: WorkbookContext,
    emitter: SseEmitter,
    loopOptions: AgenticLoopOptions = {},
  ): Promise<AgenticLoopResult> {
    return this.runInternal(originalPrompt, subtasks, context, emitter, loopOptions);
  }

  private async runInternal(
    originalPrompt: string,
    subtasks: SubTask[],
    context: WorkbookContext,
    emitter: SseEmitter,
    loopOptions: AgenticLoopOptions,
  ): Promise<AgenticLoopResult> {
    const startedAt = Date.now();
    let iterationsRun = 0;
    let timedOut = false;
    const formulaValidationLog: FormulaValidationResult[] = [];

    const ordered = this.orderByDependencies(subtasks);
    const subtaskStates: SubtaskActionState[] = ordered.map((subtask) => ({
      subtask,
      actions: [],
      completed: false,
    }));

    const waves = computeExecutionWaves(ordered);
    const completedIds = new Set<string>();

    for (const wave of waves) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        timedOut = true;
        break;
      }

      if (wave.length > 1) {
        emitter.send({
          type: 'THINKING',
          message: `Running ${wave.length} independent steps in parallel...`,
        });
      }

      // Await every sibling to settlement so a mid-wave LLM failure cannot leave
      // other in-flight fetches as unhandled rejections (Node process crash).
      const waveSettled = await Promise.all(
        wave.map(async (subtask) => {
          const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
          if (!state) {
            return { iterations: 0 as number, error: null as unknown };
          }

          const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
          try {
            const iterations = await this.executeSubtask(
              state,
              subtaskStates,
              context,
              emitter,
              startedAt,
              formulaValidationLog,
              loopOptions,
              visibleIds,
              () => {
                timedOut = true;
              },
            );
            return { iterations, error: null as unknown };
          } catch (error) {
            const reason =
              error instanceof Error ? error.message : 'Parallel subtask execution failed';
            state.failedReason = reason;
            this.logger.warn(
              `Parallel subtask "${subtask.description}" failed: ${reason}`,
            );
            return { iterations: 0, error };
          }
        }),
      );

      iterationsRun += waveSettled.reduce((sum, entry) => sum + entry.iterations, 0);
      for (const subtask of wave) {
        completedIds.add(subtask.id);
      }

      // Do not rethrow — keep going so completed siblings can surface as partialProgress.
      // Siblings were awaited above so mid-wave LLM aborts cannot become unhandled rejections.
      const failedInWave = waveSettled.filter((entry) => entry.error);
      if (failedInWave.length > 0) {
        emitter.send({
          type: 'THINKING',
          message: `${failedInWave.length} parallel step(s) failed — continuing with remaining work`,
        });
      }
    }

    if (timedOut) {
      this.logger.warn('Agentic loop timed out before completion');
      emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
      return this.buildLoopResult(subtaskStates, iterationsRun, false, {
        preferCompletedOnly: true,
        defaultFailReason: 'Agentic loop timed out before completion',
      });
    }

    let verifierPassed = false;
    let verifierCycle = 0;
    const stepRetryAttempts = new Map<string, number>();
    const maxVerifierCycles = Math.max(1, ordered.length * (this.MAX_STEP_RETRIES + 1));
    let retryExhaustedMessage: string | null = null;
    const validatorSummary = this.formulaValidator.summarizeForVerifier(formulaValidationLog);
    loopOptions.parseFailureTracker ??= { hadFailure: false };
    let lastSubtaskVerifyResults: Array<{
      subtaskId: string;
      passed: boolean;
      feedback: string;
      inconclusive?: boolean;
    }> = [];
    /** Spec 17 Bug B: once a subtask passes verification, never re-execute it. */
    const lockedPassIds = new Set<string>();

    while (!verifierPassed && verifierCycle < maxVerifierCycles && !timedOut) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        timedOut = true;
        emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
        break;
      }

      verifierCycle += 1;
      emitter.send({ type: 'THINKING', message: 'Running deterministic checks...' });

      const shadow = this.buildShadowFromStates(context, subtaskStates);
      const verifyContext = this.enrichContextFromShadow(shadow);
      const cheapChecks = this.runDeterministicChecks(
        originalPrompt,
        ordered,
        subtaskStates,
        verifyContext,
      );

      if (cheapChecks.passed && !cheapChecks.requiresLlmVerification) {
        verifierPassed = true;
        lastSubtaskVerifyResults = ordered.map((subtask) => ({
          subtaskId: subtask.id,
          passed: true,
          feedback: 'Deterministic checks passed',
        }));
        for (const subtask of ordered) lockedPassIds.add(subtask.id);
        emitter.send({ type: 'VERIFY_PASS' });
        this.logger.log('Skipped LLM verifier — deterministic checks passed cleanly');
        break;
      }

      if (!cheapChecks.passed) {
        this.logger.warn(
          `Deterministic checks failed cycle ${verifierCycle}: ${cheapChecks.feedback}`,
        );
        emitter.send({
          type: 'THINKING',
          message: `Fixing issues: ${cheapChecks.feedback}`,
        });

        lastSubtaskVerifyResults = cheapChecks.subtaskResults.map((result) => ({
          subtaskId: result.subtaskId,
          passed: result.passed,
          feedback: result.feedback,
        }));
        this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);

        const failingIds = this.collectFailingIdsForRetry(
          cheapChecks.subtaskResults,
          lockedPassIds,
        );
        if (failingIds.size === 0) {
          break;
        }

        const exhaustedIds = this.findExhaustedSubtasks(failingIds, stepRetryAttempts);
        if (exhaustedIds.size > 0) {
          retryExhaustedMessage = this.buildRetryExhaustedMessage(exhaustedIds, cheapChecks.feedback);
          for (const id of exhaustedIds) {
            const state = subtaskStates.find((entry) => entry.subtask.id === id);
            if (state) {
              state.failedReason =
                state.failedReason ??
                `Could not complete after ${this.MAX_STEP_RETRIES} attempts: ${cheapChecks.feedback}`;
            }
          }
          break;
        }

        const toRetry = ordered.filter((subtask) => failingIds.has(subtask.id));
        this.logger.log(
          `Selective retry (deterministic): re-executing [${[...failingIds].join(', ')}] — locked passes: [${[...lockedPassIds].join(', ')}]`,
        );

        iterationsRun += await this.retrySubtasks(
          toRetry,
          subtaskStates,
          context,
          emitter,
          startedAt,
          formulaValidationLog,
          loopOptions,
          failingIds,
          cheapChecks,
          () => {
            timedOut = true;
          },
          undefined,
          stepRetryAttempts,
        );
        continue;
      }

      emitter.send({ type: 'THINKING', message: 'Verifying semantic correctness...' });

      const allActions = this.flattenActions(subtaskStates);
      const hasFormulaActions = allActions.some(
        (action) => action.type === 'SET_FORMULA' || action.type === 'FILL_DOWN',
      );
      const skipDecision = shouldSkipVerifier({
        actions: allActions,
        subtaskCount: ordered.length,
        executorParsedOnFirstAttempt: !loopOptions.parseFailureTracker?.hadFailure,
        hasFormulaActions,
      });

      let verification: VerifierOutput;
      if (skipDecision.skip) {
        this.logger.log(
          `[${loopOptions.correlationId ?? '-'}] ${skipDecision.reason}`,
        );
        verification = {
          passed: true,
          feedback: skipDecision.reason,
          issues: [],
          subtaskResults: ordered.map((subtask) => ({
            subtaskId: subtask.id,
            passed: true,
            feedback: skipDecision.reason,
            issues: [],
          })),
        };
      } else {
        this.logger.log(
          `[${loopOptions.correlationId ?? '-'}] Running Verifier: ${skipDecision.reason}`,
        );
        // Spec 17: only verify subtasks that are not already locked as passed.
        const toVerify = ordered.filter((subtask) => !lockedPassIds.has(subtask.id));
        if (toVerify.length === 0) {
          verification = {
            passed: true,
            feedback: 'All subtasks already verified',
            issues: [],
            subtaskResults: ordered.map((subtask) => ({
              subtaskId: subtask.id,
              passed: true,
              feedback: 'Previously verified',
              issues: [],
            })),
          };
        } else {
          const partial = await this.verifier.verify(
            originalPrompt,
            toVerify,
            this.actionsBySubtaskMap(subtaskStates),
            verifyContext,
            validatorSummary,
            loopOptions.correlationId,
          );
          verification = this.mergeWithLockedPasses(partial, ordered, lockedPassIds, lastSubtaskVerifyResults);
        }
      }

      if (verification.passed) {
        verifierPassed = true;
        lastSubtaskVerifyResults = verification.subtaskResults.map((result) => ({
          subtaskId: result.subtaskId,
          passed: result.passed,
          feedback: result.feedback,
          inconclusive: result.inconclusive,
        }));
        this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);
        emitter.send({ type: 'VERIFY_PASS' });
        break;
      }

      this.logger.warn(`Verifier failed cycle ${verifierCycle}: ${verification.feedback}`);
      emitter.send({
        type: 'THINKING',
        message: `Fixing issues: ${verification.feedback}`,
      });

      lastSubtaskVerifyResults = verification.subtaskResults.map((result) => ({
        subtaskId: result.subtaskId,
        passed: result.passed,
        feedback: result.feedback,
        inconclusive: result.inconclusive,
      }));
      this.lockPassedSubtasks(lastSubtaskVerifyResults, lockedPassIds);

      const failingIds = this.collectFailingIdsForRetry(
        verification.subtaskResults,
        lockedPassIds,
      );

      // Inconclusive-only: re-verify without re-executing (next loop cycle).
      const inconclusiveOnly =
        failingIds.size === 0 &&
        verification.subtaskResults.some((r) => r.inconclusive && !lockedPassIds.has(r.subtaskId));
      if (inconclusiveOnly) {
        this.logger.warn(
          'Verifier had inconclusive (truncated) results — re-verifying without re-execution',
        );
        continue;
      }

      if (failingIds.size === 0) {
        break;
      }

      const exhaustedIds = this.findExhaustedSubtasks(failingIds, stepRetryAttempts);
      if (exhaustedIds.size > 0) {
        retryExhaustedMessage = this.buildRetryExhaustedMessage(exhaustedIds, verification.feedback);
        for (const id of exhaustedIds) {
          const state = subtaskStates.find((entry) => entry.subtask.id === id);
          if (state) {
            state.failedReason =
              state.failedReason ??
              `Could not complete after ${this.MAX_STEP_RETRIES} attempts: ${verification.feedback}`;
          }
        }
        break;
      }

      const toRetry = ordered.filter((subtask) => failingIds.has(subtask.id));
      this.logger.log(
        `Selective retry: re-executing [${[...failingIds].join(', ')}] — locked passes: [${[...lockedPassIds].join(', ')}]`,
      );

      iterationsRun += await this.retrySubtasks(
        toRetry,
        subtaskStates,
        context,
        emitter,
        startedAt,
        formulaValidationLog,
        loopOptions,
        failingIds,
        {
          passed: false,
          requiresLlmVerification: true,
          feedback: verification.feedback,
          issues: verification.issues,
          subtaskResults: verification.subtaskResults,
        },
        () => {
          timedOut = true;
        },
        (subtask) => {
          const subtaskResult = verification.subtaskResults.find(
            (result) => result.subtaskId === subtask.id,
          );
          return {
            verifierFeedback: subtaskResult?.feedback ?? verification.feedback,
            verifierIssues: subtaskResult?.issues ?? verification.issues,
          };
        },
        stepRetryAttempts,
      );
    }

    if (!verifierPassed) {
      emitter.send({
        type: 'VERIFY_FAIL',
        feedback:
          retryExhaustedMessage ?? 'Could not verify after scoped retries — showing best attempt',
      });
    }

    this.applyVerifyResultsToStates(subtaskStates, lastSubtaskVerifyResults, verifierPassed);

    return this.buildLoopResult(subtaskStates, iterationsRun, verifierPassed, {
      preferCompletedOnly: !verifierPassed,
      defaultFailReason:
        retryExhaustedMessage ?? 'Could not complete and verify the full request',
      timedOut,
    });
  }

  private applyVerifyResultsToStates(
    subtaskStates: SubtaskActionState[],
    results: Array<{ subtaskId: string; passed: boolean; feedback: string }>,
    verifierPassed: boolean,
  ): void {
    if (verifierPassed) {
      for (const state of subtaskStates) {
        state.verified = true;
        state.completed = true;
      }
      return;
    }

    for (const result of results) {
      const state = subtaskStates.find((entry) => entry.subtask.id === result.subtaskId);
      if (!state) continue;
      state.verified = result.passed;
      if (result.passed) {
        state.completed = true;
      } else if (!state.failedReason) {
        state.failedReason = result.feedback || 'Verification failed for this step';
      }
    }
  }

  private buildLoopResult(
    subtaskStates: SubtaskActionState[],
    iterationsRun: number,
    verifierPassed: boolean,
    options: {
      preferCompletedOnly: boolean;
      defaultFailReason: string;
      timedOut?: boolean;
    },
  ): AgenticLoopResult {
    const completedSubtasks: CompletedSubtaskResult[] = subtaskStates
      .filter((state) => {
        if (!state.completed || state.actions.length === 0) return false;
        if (verifierPassed) return true;
        // Partial progress: only include steps that verified or finished cleanly before a later failure
        return state.verified === true || (state.completed && !state.failedReason);
      })
      .map((state) => ({
        subtaskId: state.subtask.id,
        actions: state.actions,
        verified: state.verified === true || verifierPassed,
      }));

    const failedState =
      subtaskStates.find((state) => Boolean(state.failedReason)) ??
      (!verifierPassed
        ? subtaskStates.find((state) => state.verified === false) ??
          subtaskStates.find((state) => !state.completed)
        : undefined);

    const failedSubtask: FailedSubtaskResult | null =
      !verifierPassed && failedState
        ? {
            subtaskId: failedState.subtask.id,
            reason:
              failedState.failedReason ??
              (options.timedOut
                ? 'Timed out before this step completed'
                : options.defaultFailReason),
          }
        : null;

    const partialProgress =
      !verifierPassed && completedSubtasks.length > 0 && failedSubtask !== null;

    const actions =
      options.preferCompletedOnly && partialProgress
        ? completedSubtasks.flatMap((entry) => entry.actions)
        : this.flattenActions(subtaskStates);

    return {
      actions,
      iterationsRun,
      verifierPassed,
      completedSubtasks,
      failedSubtask,
      partialProgress,
    };
  }

  private runDeterministicChecks(
    originalPrompt: string,
    subtasks: SubTask[],
    subtaskStates: SubtaskActionState[],
    context: WorkbookContext,
  ): CheckerResult {
    const completeness = this.completenessChecker.check(subtasks, subtaskStates);
    const formatting = this.formattingChecker.check(subtaskStates, context);
    const semantic = this.semanticFormulaChecker.check(
      originalPrompt,
      subtasks,
      subtaskStates,
      context,
    );
    const merged = mergeCheckerResults([completeness, formatting, semantic]);

    const needsSemanticReview =
      subtaskStates.some((state) =>
        state.actions.some(
          (action) =>
            action.type === 'SET_FORMULA' ||
            (typeof action.formula === 'string' && action.formula.startsWith('=')),
        ),
      ) ||
      subtasks.some((subtask) => subtask.dependsOn.length > 0) ||
      subtasks.length > 2;

    return {
      ...merged,
      requiresLlmVerification: merged.requiresLlmVerification || needsSemanticReview,
    };
  }

  private async retrySubtasks(
    subtasksToRetry: SubTask[],
    subtaskStates: SubtaskActionState[],
    context: WorkbookContext,
    emitter: SseEmitter,
    startedAt: number,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    failingIds: Set<string>,
    checkResult: CheckerResult,
    onTimeout: () => void,
    retryContextFor?: (subtask: SubTask) => RetryContext,
    stepRetryAttempts?: Map<string, number>,
  ): Promise<number> {
    let iterationsRun = 0;
    const completedIds = new Set(
      subtaskStates
        .map((state) => state.subtask.id)
        .filter((id) => !failingIds.has(id)),
    );

    for (const subtask of subtasksToRetry) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        break;
      }

      const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
      if (!state) continue;

      const subtaskResult = checkResult.subtaskResults.find(
        (result) => result.subtaskId === subtask.id,
      );
      const feedback = subtaskResult?.feedback ?? checkResult.feedback;
      const issues: VerifierIssue[] = subtaskResult?.issues ?? checkResult.issues;

      state.actions = [];
      state.completed = false;
      state.verified = undefined;
      state.failedReason = undefined;
      const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
      const retryAttempt = (stepRetryAttempts?.get(subtask.id) ?? 0) + 1;
      stepRetryAttempts?.set(subtask.id, retryAttempt);
      iterationsRun += await this.executeSubtask(
        state,
        subtaskStates,
        context,
        emitter,
        startedAt,
        formulaValidationLog,
        loopOptions,
        visibleIds,
        onTimeout,
        retryContextFor?.(subtask) ?? { verifierFeedback: feedback, verifierIssues: issues },
        retryAttempt,
      );
    }

    return iterationsRun;
  }

  private async executeSubtask(
    state: SubtaskActionState,
    allStates: SubtaskActionState[],
    baseContext: WorkbookContext,
    emitter: SseEmitter,
    startedAt: number,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    visibleStateIds: Set<string>,
    onTimeout: () => void,
    retryContext?: RetryContext,
    retryAttempt?: number,
  ): Promise<number> {
    const { subtask } = state;
    let iterationsRun = 0;

    this.logger.log(`Agentic loop: starting subtask "${subtask.description}"`);
    emitter.send({ type: 'CHECKPOINT', step: subtask.description });

    let iteration = 0;
    let subtaskDone = false;

    while (!subtaskDone && iteration < this.MAX_ITERATIONS_PER_SUBTASK) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        break;
      }

      iteration += 1;
      iterationsRun += 1;

      emitter.send({
        type: 'THINKING',
        message:
          iteration > 1
            ? `Continuing "${subtask.description}" (step ${iteration})...`
            : `Working on "${subtask.description}"...`,
      });

      const shadow = this.buildShadowFromStates(baseContext, allStates, visibleStateIds);
      const currentContext = {
        ...this.enrichContextFromShadow(shadow),
        ...retryContext,
      };
      const previousActions = this.flattenActions(
        allStates.filter((entry) => visibleStateIds.has(entry.subtask.id)),
      );

      const validatedBatch = await this.runExecutorWithFormulaValidation(
        subtask,
        currentContext,
        shadow,
        previousActions,
        emitter,
        startedAt,
        onTimeout,
        formulaValidationLog,
        loopOptions,
        retryAttempt && retryContext?.verifierFeedback
          ? {
              originalStep: subtask,
              attempt: retryAttempt,
              maxAttempts: this.MAX_STEP_RETRIES,
              verifierFeedback: retryContext.verifierFeedback,
            }
          : undefined,
      );

      if (!validatedBatch) {
        emitter.send({
          type: 'THINKING',
          message: `Formula validation blocked actions for "${subtask.description}" — retrying`,
        });
        continue;
      }

      for (const action of validatedBatch.actions) {
        emitter.send({ type: 'ACTION', action });
        state.actions.push(action);
      }

      subtaskDone = validatedBatch.isDone;

      if (
        !subtaskDone &&
        validatedBatch.actions.length === 0 &&
        isExecutorBlockedSignal(validatedBatch.nextStep)
      ) {
        state.failedReason = validatedBatch.nextStep;
        this.logger.warn(
          `Executor blocked on "${subtask.description}": ${validatedBatch.nextStep}`,
        );
        emitter.send({
          type: 'THINKING',
          message: validatedBatch.nextStep ?? 'Blocked — cannot complete this step',
        });
        break;
      }

      if (!subtaskDone && validatedBatch.nextStep) {
        emitter.send({ type: 'THINKING', message: validatedBatch.nextStep });
      }
    }

    if (subtaskDone) {
      state.completed = true;
      state.failedReason = undefined;
    } else {
      const hitTimeout = Date.now() - startedAt > this.TIMEOUT_MS;
      const reason = hitTimeout
        ? `Subtask "${subtask.description}" timed out before completion`
        : `Subtask "${subtask.description}" hit max iterations (${this.MAX_ITERATIONS_PER_SUBTASK})`;
      state.failedReason = reason;
      this.logger.warn(reason);
      emitter.send({
        type: 'THINKING',
        message: `Reached step limit for "${subtask.description}" — moving on`,
      });
    }

    return iterationsRun;
  }

  /** Run executor with optional tool fetch, pre/post formula validation, and retries. */
  private async runExecutorWithFormulaValidation(
    subtask: SubTask,
    context: WorkbookContext,
    baseShadow: ShadowWorkbook,
    previousActions: Action[],
    emitter: SseEmitter,
    startedAt: number,
    onTimeout: () => void,
    formulaValidationLog: FormulaValidationResult[],
    loopOptions: AgenticLoopOptions,
    retryStepContext?: StepRetryContext,
  ): Promise<ExecutorOutput | null> {
    const callStartedAt = Date.now();
    const model = this.executor.modelName;
    let execContext = { ...context };
    let result = buildDeterministicSubtaskActions(subtask, execContext);
    if (result) {
      this.logger.log(
        `Using ${result.actions.length} deterministic action(s) for "${subtask.description}"`,
      );
    } else {
      try {
        result = retryStepContext
          ? await this.executor.retryStep(
              retryStepContext,
              execContext,
              previousActions,
              loopOptions.correlationId,
            )
          : await this.executor.execute(
              subtask,
              execContext,
              previousActions,
              loopOptions.correlationId,
            );
        this.noteExecutorParseResult(result, loopOptions);
      } catch (error) {
        if (error instanceof StepRetryExhaustedError) {
          this.logger.error(error.message);
          this.structuredLogger.logAgentEvent({
            correlationId: loopOptions.correlationId ?? '-',
            agent: 'workbook',
            model,
            durationMs: Date.now() - callStartedAt,
            success: false,
            error: error.message,
          });
          return null;
        }
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: error instanceof Error ? error.message : String(error),
        });
        throw error;
      }
    }

    let toolAttempts = 0;
    while (
      result.toolRequest &&
      toolAttempts < this.MAX_TOOL_REQUESTS &&
      loopOptions.conversationId &&
      loopOptions.toolEmit
    ) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        return null;
      }

      toolAttempts += 1;
      const { toolRequest } = result;
      emitter.send({
        type: 'THINKING',
        message: `Fetching range ${toolRequest.range} on ${toolRequest.sheet}...`,
      });

      try {
        const fetched = await this.toolBridge.waitForRangeData(
          loopOptions.conversationId,
          toolRequest,
          loopOptions.toolEmit,
        );

        if (fetched.error) {
          emitter.send({
            type: 'THINKING',
            message: `Range fetch failed: ${fetched.error}`,
          });
          return null;
        }

        execContext = this.mergeFetchedRange(execContext, toolRequest, fetched.values);
        result = await this.executor.execute(
          subtask,
          execContext,
          previousActions,
          loopOptions.correlationId,
        );
        this.noteExecutorParseResult(result, loopOptions);
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Range fetch failed';
        this.logger.warn(`Tool request failed: ${message}`);
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: message,
        });
        emitter.send({ type: 'THINKING', message });
        return null;
      }
    }

    if (result.toolRequest && result.actions.length === 0) {
      emitter.send({
        type: 'THINKING',
        message: 'Could not fetch requested range — try a smaller range or simpler task',
      });
      return null;
    }

    let formulaAttempt = 0;

    while (formulaAttempt <= this.MAX_FORMULA_RETRIES) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        onTimeout();
        return null;
      }

      const preValidation = this.formulaValidator.validatePreApply(
        result.actions,
        execContext,
        subtask.targetSheet,
      );
      formulaValidationLog.push(preValidation);

      if (!preValidation.passed) {
      if (formulaAttempt >= this.MAX_FORMULA_RETRIES) {
        const deterministic = buildDeterministicSubtaskActions(subtask, execContext);
        if (deterministic?.actions.length) {
          this.logger.log(
            `Formula pre-validation failed — using ${deterministic.actions.length} deterministic action(s)`,
          );
          result = deterministic;
          formulaAttempt = 0;
          continue;
        }
        this.logger.warn(
          `Formula pre-validation failed after ${this.MAX_FORMULA_RETRIES} retries for "${subtask.description}"`,
        );
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: this.formulaValidator.formatFeedback(preValidation.issues),
        });
        emitter.send({
          type: 'THINKING',
          message: `Formula errors remain: ${this.formulaValidator.formatFeedback(preValidation.issues)}`,
        });
        return null;
      }

        formulaAttempt += 1;
        execContext = {
          ...execContext,
          formulaValidationFeedback: this.formulaValidator.formatFeedback(preValidation.issues),
          formulaValidationIssues: preValidation.issues,
        };
        result = await this.executor.execute(
          subtask,
          execContext,
          previousActions,
          loopOptions.correlationId,
        );
        this.noteExecutorParseResult(result, loopOptions);
        continue;
      }

      if (result.actions.length === 0) {
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: true,
          parsedResponse: result,
        });
        return result;
      }

      const postShadow = virtualApply(baseShadow, result.actions);
      const postValidation = this.formulaValidator.checkPostApply(
        postShadow,
        result.actions,
        execContext,
        subtask.targetSheet,
      );
      formulaValidationLog.push(postValidation);

      if (postValidation.passed) {
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: true,
          parsedResponse: result,
        });
        return result;
      }

      if (formulaAttempt >= this.MAX_FORMULA_RETRIES) {
        const deterministic = buildDeterministicSubtaskActions(subtask, execContext);
        if (deterministic?.actions.length) {
          this.logger.log(
            `Formula post-validation failed — using ${deterministic.actions.length} deterministic action(s)`,
          );
          result = deterministic;
          formulaAttempt = 0;
          continue;
        }
        this.logger.warn(
          `Formula post-validation failed after ${this.MAX_FORMULA_RETRIES} retries for "${subtask.description}"`,
        );
        this.structuredLogger.logAgentEvent({
          correlationId: loopOptions.correlationId ?? '-',
          agent: 'workbook',
          model,
          durationMs: Date.now() - callStartedAt,
          success: false,
          error: this.formulaValidator.formatFeedback(postValidation.issues),
        });
        emitter.send({
          type: 'THINKING',
          message: `Post-apply formula errors: ${this.formulaValidator.formatFeedback(postValidation.issues)}`,
        });
        return null;
      }

      formulaAttempt += 1;
      execContext = {
        ...execContext,
        formulaValidationFeedback: this.formulaValidator.formatFeedback(postValidation.issues),
        formulaValidationIssues: postValidation.issues,
      };
      result = await this.executor.execute(
        subtask,
        execContext,
        previousActions,
        loopOptions.correlationId,
      );
    }

    this.structuredLogger.logAgentEvent({
      correlationId: loopOptions.correlationId ?? '-',
      agent: 'workbook',
      model,
      durationMs: Date.now() - callStartedAt,
      success: false,
      error: `Executor returned null for subtask ${subtask.id}`,
    });
    return null;
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

    const fetchedRanges = [
      ...(context.fetchedRanges ?? []),
      {
        sheet: toolRequest.sheet,
        range: toolRequest.range,
        rowCount: values.length,
      },
    ];

    return {
      ...context,
      sheets,
      fetchedRanges,
    };
  }

  private buildShadowFromStates(
    baseContext: WorkbookContext,
    states: SubtaskActionState[],
    visibleStateIds?: Set<string>,
  ): ShadowWorkbook {
    let shadow = buildShadowWorkbook(baseContext);
    for (const state of states) {
      if (visibleStateIds && !visibleStateIds.has(state.subtask.id)) continue;
      if (state.actions.length === 0) continue;
      shadow = virtualApply(shadow, state.actions);
    }
    return shadow;
  }

  private enrichContextFromShadow(shadow: ShadowWorkbook): WorkbookContext {
    const context = shadowToWorkbookContext(shadow);
    return {
      ...context,
      sheets: context.sheets.map((sheet) => ({
        ...sheet,
        formulaInsights: this.formulaAnalyzer.analyzeSheet(sheet),
      })),
    };
  }

  private flattenActions(states: SubtaskActionState[]): Action[] {
    return states.flatMap((state) => state.actions);
  }

  private actionsBySubtaskMap(states: SubtaskActionState[]): Record<string, Action[]> {
    return Object.fromEntries(states.map((state) => [state.subtask.id, state.actions]));
  }

  private collectDownstreamSubtasks(
    ordered: SubTask[],
    failingIds: Set<string>,
  ): Set<string> {
    const downstream = new Set<string>(failingIds);

    let changed = true;
    while (changed) {
      changed = false;
      for (const subtask of ordered) {
        if (downstream.has(subtask.id)) continue;
        if (subtask.dependsOn.some((dep) => downstream.has(dep))) {
          downstream.add(subtask.id);
          changed = true;
        }
      }
    }

    return downstream;
  }

  private lockPassedSubtasks(
    results: Array<{ subtaskId: string; passed: boolean; inconclusive?: boolean }>,
    lockedPassIds: Set<string>,
  ): void {
    for (const result of results) {
      if (result.passed && !result.inconclusive) {
        lockedPassIds.add(result.subtaskId);
      }
    }
  }

  /**
   * Spec 17 Bug B: only genuinely failed (not inconclusive, not already locked)
   * subtasks are re-executed.
   */
  private collectFailingIdsForRetry(
    results: Array<{ subtaskId: string; passed: boolean; inconclusive?: boolean }>,
    lockedPassIds: Set<string>,
  ): Set<string> {
    const failing = new Set<string>();
    for (const result of results) {
      if (lockedPassIds.has(result.subtaskId)) continue;
      if (result.inconclusive) continue;
      if (!result.passed) failing.add(result.subtaskId);
    }
    return failing;
  }

  private mergeWithLockedPasses(
    partial: VerifierOutput,
    ordered: SubTask[],
    lockedPassIds: Set<string>,
    previous: Array<{ subtaskId: string; passed: boolean; feedback: string; inconclusive?: boolean }>,
  ): VerifierOutput {
    const byId = new Map(partial.subtaskResults.map((r) => [r.subtaskId, r]));
    const subtaskResults = ordered.map((subtask) => {
      if (lockedPassIds.has(subtask.id)) {
        const prev = previous.find((r) => r.subtaskId === subtask.id);
        return {
          subtaskId: subtask.id,
          passed: true,
          feedback: prev?.feedback ?? 'Previously verified',
          issues: [],
        };
      }
      return (
        byId.get(subtask.id) ?? {
          subtaskId: subtask.id,
          passed: false,
          feedback: 'Missing verifier result',
          issues: [],
          inconclusive: true,
        }
      );
    });

    const passed = subtaskResults.every((r) => r.passed && !r.inconclusive);
    return {
      passed,
      feedback: partial.feedback,
      issues: partial.issues,
      subtaskResults,
    };
  }

  private findExhaustedSubtasks(
    failingIds: Set<string>,
    stepRetryAttempts: Map<string, number>,
  ): Set<string> {
    const exhausted = new Set<string>();
    for (const subtaskId of failingIds) {
      if ((stepRetryAttempts.get(subtaskId) ?? 0) >= this.MAX_STEP_RETRIES) {
        exhausted.add(subtaskId);
      }
    }
    return exhausted;
  }

  private buildRetryExhaustedMessage(exhaustedIds: Set<string>, feedback: string): string {
    const ids = Array.from(exhaustedIds).join(', ');
    return `Could not complete step(s) ${ids} after ${this.MAX_STEP_RETRIES} attempts. ${feedback}`;
  }

  private noteExecutorParseResult(
    result: ExecutorOutput | null | undefined,
    loopOptions: AgenticLoopOptions,
  ): void {
    if (result?.parsedOnFirstAttempt === false && loopOptions.parseFailureTracker) {
      loopOptions.parseFailureTracker.hadFailure = true;
    }
  }

  private orderByDependencies(subtasks: SubTask[]): SubTask[] {
    const byId = new Map(subtasks.map((subtask) => [subtask.id, subtask]));
    const visited = new Set<string>();
    const result: SubTask[] = [];

    const visit = (id: string) => {
      if (visited.has(id)) return;
      const task = byId.get(id);
      if (!task) return;
      for (const dep of task.dependsOn) {
        visit(dep);
      }
      visited.add(id);
      result.push(task);
    };

    for (const task of subtasks) {
      visit(task.id);
    }

    return result;
  }
}
