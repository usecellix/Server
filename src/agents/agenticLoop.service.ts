import { Injectable, Logger } from '@nestjs/common';
import { ExecutorAgent } from './executor.agent';
import { VerifierAgent } from './verifier.agent';
import {
  Action,
  ExecutorOutput,
  SubTask,
  VerifierIssue,
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
import { CheckerResult, mergeCheckerResults } from './checkers/checker.types';
import { buildDeterministicSubtaskActions } from './utils/compound-action.util';

export interface AgenticLoopOptions {
  conversationId?: string;
  toolEmit?: (event: string, data: Record<string, unknown>) => void;
}

export interface AgenticLoopResult {
  actions: Action[];
  iterationsRun: number;
  verifierPassed: boolean;
}

interface SubtaskActionState {
  subtask: SubTask;
  actions: Action[];
}

type RetryContext = Pick<
  WorkbookContext,
  'verifierFeedback' | 'verifierIssues' | 'formulaValidationFeedback' | 'formulaValidationIssues'
>;

@Injectable()
export class AgenticLoopService {
  private readonly logger = new Logger(AgenticLoopService.name);
  private readonly MAX_ITERATIONS_PER_SUBTASK = 10;
  private readonly MAX_VERIFY_RETRIES = 2;
  private readonly MAX_FORMULA_RETRIES = 2;
  private readonly MAX_TOOL_REQUESTS = 5;
  private readonly TIMEOUT_MS = 180_000;

  constructor(
    private readonly executor: ExecutorAgent,
    private readonly verifier: VerifierAgent,
    private readonly formulaAnalyzer: FormulaAnalyzer,
    private readonly formulaValidator: FormulaValidatorService,
    private readonly toolBridge: ToolBridgeService,
    private readonly completenessChecker: CompletenessChecker,
    private readonly formattingChecker: FormattingChecker,
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

      const waveResults = await Promise.all(
        wave.map(async (subtask) => {
          const state = subtaskStates.find((entry) => entry.subtask.id === subtask.id);
          if (!state) return 0;

          const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
          return this.executeSubtask(
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
        }),
      );

      iterationsRun += waveResults.reduce((sum, count) => sum + count, 0);
      for (const subtask of wave) {
        completedIds.add(subtask.id);
      }
    }

    if (timedOut) {
      this.logger.warn('Agentic loop timed out before completion');
      emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
      return {
        actions: this.flattenActions(subtaskStates),
        iterationsRun,
        verifierPassed: false,
      };
    }

    let verifierPassed = false;
    let verifyAttempt = 0;
    const validatorSummary = this.formulaValidator.summarizeForVerifier(formulaValidationLog);

    while (!verifierPassed && verifyAttempt < this.MAX_VERIFY_RETRIES && !timedOut) {
      if (Date.now() - startedAt > this.TIMEOUT_MS) {
        timedOut = true;
        emitter.send({ type: 'ERROR', message: 'Agentic loop timeout' });
        break;
      }

      verifyAttempt += 1;
      emitter.send({ type: 'THINKING', message: 'Running deterministic checks...' });

      const shadow = this.buildShadowFromStates(context, subtaskStates);
      const verifyContext = this.enrichContextFromShadow(shadow);
      const cheapChecks = this.runDeterministicChecks(ordered, subtaskStates, verifyContext);

      if (cheapChecks.passed && !cheapChecks.requiresLlmVerification) {
        verifierPassed = true;
        emitter.send({ type: 'VERIFY_PASS' });
        this.logger.log('Skipped LLM verifier — deterministic checks passed cleanly');
        break;
      }

      if (!cheapChecks.passed) {
        this.logger.warn(
          `Deterministic checks failed attempt ${verifyAttempt}: ${cheapChecks.feedback}`,
        );
        emitter.send({
          type: 'THINKING',
          message: `Fixing issues: ${cheapChecks.feedback}`,
        });

        if (verifyAttempt >= this.MAX_VERIFY_RETRIES) {
          break;
        }

        const failingIds = new Set(
          cheapChecks.subtaskResults.filter((result) => !result.passed).map((result) => result.subtaskId),
        );
        if (failingIds.size === 0) {
          break;
        }

        const downstreamIds = this.collectDownstreamSubtasks(ordered, failingIds);
        const toRetry = ordered.filter((subtask) => downstreamIds.has(subtask.id));

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
        );
        continue;
      }

      emitter.send({ type: 'THINKING', message: 'Verifying semantic correctness...' });
      const verification = await this.verifier.verify(
        originalPrompt,
        ordered,
        this.actionsBySubtaskMap(subtaskStates),
        verifyContext,
        validatorSummary,
      );

      if (verification.passed) {
        verifierPassed = true;
        emitter.send({ type: 'VERIFY_PASS' });
        break;
      }

      this.logger.warn(`Verifier failed attempt ${verifyAttempt}: ${verification.feedback}`);
      emitter.send({
        type: 'THINKING',
        message: `Fixing issues: ${verification.feedback}`,
      });

      if (verifyAttempt >= this.MAX_VERIFY_RETRIES) {
        break;
      }

      const failingIds = new Set(
        verification.subtaskResults.filter((r) => !r.passed).map((r) => r.subtaskId),
      );

      if (failingIds.size === 0) {
        break;
      }

      const downstreamIds = this.collectDownstreamSubtasks(ordered, failingIds);
      const toRetry = ordered.filter((subtask) => downstreamIds.has(subtask.id));

      iterationsRun += await this.retrySubtasks(
        toRetry,
        subtaskStates,
        context,
        emitter,
        startedAt,
        formulaValidationLog,
        loopOptions,
        downstreamIds,
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
      );
    }

    if (!verifierPassed) {
      emitter.send({
        type: 'VERIFY_FAIL',
        feedback: 'Could not verify after retries — showing best attempt',
      });
    }

    return {
      actions: this.flattenActions(subtaskStates),
      iterationsRun,
      verifierPassed,
    };
  }

  private runDeterministicChecks(
    subtasks: SubTask[],
    subtaskStates: SubtaskActionState[],
    context: WorkbookContext,
  ): CheckerResult {
    const completeness = this.completenessChecker.check(subtasks, subtaskStates);
    const formatting = this.formattingChecker.check(subtaskStates, context);
    const merged = mergeCheckerResults([completeness, formatting]);

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
      const visibleIds = new Set([...completedIds, ...subtask.dependsOn]);
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

      if (!subtaskDone && validatedBatch.nextStep) {
        emitter.send({ type: 'THINKING', message: validatedBatch.nextStep });
      }
    }

    if (!subtaskDone) {
      const hitTimeout = Date.now() - startedAt > this.TIMEOUT_MS;
      this.logger.warn(
        hitTimeout
          ? `Subtask "${subtask.description}" timed out before completion`
          : `Subtask "${subtask.description}" hit max iterations (${this.MAX_ITERATIONS_PER_SUBTASK})`,
      );
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
  ): Promise<ExecutorOutput | null> {
    let execContext = { ...context };
    let result = buildDeterministicSubtaskActions(subtask, execContext);
    if (result) {
      this.logger.log(
        `Using ${result.actions.length} deterministic action(s) for "${subtask.description}"`,
      );
    } else {
      result = await this.executor.execute(subtask, execContext, previousActions);
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
        result = await this.executor.execute(subtask, execContext, previousActions);
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Range fetch failed';
        this.logger.warn(`Tool request failed: ${message}`);
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
        result = await this.executor.execute(subtask, execContext, previousActions);
        continue;
      }

      if (result.actions.length === 0) {
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
      result = await this.executor.execute(subtask, execContext, previousActions);
    }

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
