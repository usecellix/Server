import { Injectable, Logger } from '@nestjs/common';
import { AppConfigService } from '../config/app-config.service';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import { PLANNER_SYSTEM_PROMPT, buildPlannerUserMessage } from './prompts/planner.prompt';
import { parseAgentJson } from './utils/parse-agent-json.util';
import { buildCompoundFallbackSubtasks } from './utils/compound-action.util';
import { PlannerOutput, SubTask, WorkbookContext } from './types/agent.types';

const JSON_RETRY_SUFFIX =
  '\n\nIMPORTANT: Your previous response was not valid JSON. Reply with ONLY a single JSON object matching the schema — no markdown fences, no commentary.';

@Injectable()
export class PlannerAgent {
  private readonly logger = new Logger(PlannerAgent.name);

  constructor(
    private readonly llm: OpenRouterService,
    private readonly config: AppConfigService,
  ) {}

  async plan(
    prompt: string,
    context: WorkbookContext,
    history: { role: string; content: string }[] = [],
    promptContext?: string,
  ): Promise<PlannerOutput> {
    const userMessage = buildPlannerUserMessage(prompt, context, history, promptContext);

    let raw = await this.llm.complete({
      systemPrompt: PLANNER_SYSTEM_PROMPT,
      userMessage,
      model: this.config.openRouterModelHigh,
      temperature: 0.2,
      maxTokens: 1000,
    });

    let parsed = this.tryParsePlanner(raw);
    if (!parsed) {
      this.logger.warn('Planner JSON parse failed — retrying once');
      raw = await this.llm.complete({
        systemPrompt: PLANNER_SYSTEM_PROMPT,
        userMessage: userMessage + JSON_RETRY_SUFFIX,
        model: this.config.openRouterModelHigh,
        temperature: 0.1,
        maxTokens: 1000,
      });
      parsed = this.tryParsePlanner(raw);
    }

    if (parsed) {
      this.logger.log(
        `Planner produced ${parsed.subtasks.length} subtasks, confidence: ${parsed.confidence}`,
      );
      return parsed;
    }

    this.logger.warn(
      `Planner returned invalid JSON after retry — using fallback plan. Raw snippet: ${this.clip(raw)}`,
    );
    return this.buildFallbackPlan(prompt, context);
  }

  private tryParsePlanner(raw: string): PlannerOutput | null {
    try {
      const parsed = parseAgentJson<Partial<PlannerOutput>>(raw);
      return this.normalizePlannerOutput(parsed);
    } catch {
      return null;
    }
  }

  private normalizePlannerOutput(parsed: Partial<PlannerOutput>): PlannerOutput {
    const subtasks = Array.isArray(parsed.subtasks)
      ? parsed.subtasks
          .filter((s): s is SubTask => Boolean(s?.id && s?.description && s?.targetSheet))
          .map((s) => ({
            id: String(s.id),
            description: String(s.description),
            targetSheet: String(s.targetSheet),
            dependsOn: Array.isArray(s.dependsOn) ? s.dependsOn.map(String) : [],
            estimatedActions:
              typeof s.estimatedActions === 'number' ? s.estimatedActions : 1,
          }))
      : [];

    const clarificationsNeeded = Array.isArray(parsed.clarificationsNeeded)
      ? parsed.clarificationsNeeded.map(String).filter(Boolean)
      : [];

    const confidence =
      parsed.confidence === 'high' ||
      parsed.confidence === 'medium' ||
      parsed.confidence === 'low'
        ? parsed.confidence
        : 'medium';

    return {
      subtasks,
      clarificationsNeeded,
      confidence,
      reasoning: String(parsed.reasoning ?? ''),
    };
  }

  private buildFallbackPlan(prompt: string, context: WorkbookContext): PlannerOutput {
    const activeSheet = context.activeSheetName || 'Sheet1';
    const sheet = context.sheets.find((s) => s.name === activeSheet);
    const hasValues = sheet?.values.some((row) =>
      row?.some((cell) => cell !== null && cell !== '' && String(cell).trim() !== ''),
    );
    const hasHeaders = (sheet?.values[0] ?? []).some(
      (cell) => cell !== null && cell !== '' && String(cell).trim() !== '',
    );
    const isEmpty = !sheet || (!hasValues && !hasHeaders && sheet.rowCount === 0);

    if (isEmpty) {
      return {
        subtasks: [],
        clarificationsNeeded: [
          'I could not read sheet data (the workbook may be empty or still loading). Please try again after the sheet has data, or specify which column to sort by.',
        ],
        confidence: 'low',
        reasoning: 'Fallback — empty or unreadable workbook context',
      };
    }

    const compoundSubtasks = buildCompoundFallbackSubtasks(prompt, context);
    if (compoundSubtasks) {
      return {
        subtasks: compoundSubtasks,
        clarificationsNeeded: [],
        confidence: 'low',
        reasoning: 'Fallback compound plan — create sheet then sort',
      };
    }

    return {
      subtasks: [
        {
          id: 's1',
          description: prompt,
          targetSheet: activeSheet,
          dependsOn: [],
          estimatedActions: 3,
        },
      ],
      clarificationsNeeded: [],
      confidence: 'low',
      reasoning: 'Fallback single-step plan — planner JSON was not parseable',
    };
  }

  private clip(value: string, max = 400): string {
    const normalized = value.replace(/\s+/g, ' ').trim();
    return normalized.length <= max ? normalized : `${normalized.slice(0, max)}...`;
  }
}
