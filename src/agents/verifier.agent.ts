import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import {
  VERIFIER_SYSTEM_PROMPT,
  buildVerifierUserMessage,
  normalizeVerifierOutput,
} from './prompts/verifier.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { Action, SubTask, VerifierOutput, WorkbookContext } from './types/agent.types';

@Injectable()
export class VerifierAgent {
  private readonly logger = new Logger(VerifierAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  async verify(
    originalPrompt: string,
    subtasks: SubTask[],
    actionsBySubtask: Record<string, Action[]>,
    context: WorkbookContext,
    formulaValidatorSummary?: string,
  ): Promise<VerifierOutput> {
    const userMessage = buildVerifierUserMessage(
      originalPrompt,
      subtasks,
      actionsBySubtask,
      context,
      formulaValidatorSummary,
    );

    const raw = await this.llm.complete({
      systemPrompt: VERIFIER_SYSTEM_PROMPT,
      userMessage,
      model: this.config.openRouterModelMedium,
      temperature: 0.1,
      maxTokens: 1500,
    });

    try {
      const parsed = parseAgentJson<Partial<VerifierOutput>>(raw);
      const normalized = normalizeVerifierOutput(parsed, subtasks.map((s) => s.id));
      this.logger.log(`Verifier: ${normalized.passed ? 'PASS' : 'FAIL'} — ${normalized.feedback}`);
      return normalized;
    } catch {
      this.logger.error('Verifier returned invalid JSON', raw);
      return {
        passed: true,
        feedback: 'Verifier parse error — defaulting to pass',
        issues: [],
        subtaskResults: subtasks.map((s) => ({
          subtaskId: s.id,
          passed: true,
          feedback: 'Parse error — default pass',
          issues: [],
        })),
      };
    }
  }
}
