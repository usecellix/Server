import { Injectable, Logger } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Conversation, ConversationDocument } from '../excel-ai/schemas/conversation.schema';

/**
 * `Conversation` is declared with `timestamps: true` but does not re-declare
 * `createdAt`/`updatedAt` as `@Prop`s, so they exist at runtime without being
 * on the class type. This local alias adds them for read purposes only —
 * preferable to editing the shared schema, which the Excel pipeline also uses.
 */
type StoredConversation = ConversationDocument & { updatedAt?: Date };
import { randomUUID } from 'node:crypto';
import { OpenRouterService } from '../excel-ai/services/openrouter.service';
import { UsageBillingService } from '../credit/usage-billing.service';
import { CreditActionType } from '../credit/types/credit.types';
import { resolveCreditCost } from '../credit/credit-cost-catalog';
import { InsufficientCreditError } from '../credit/errors/insufficient-credit.error';
import { runWithLlmUsageContext, type LlmUsageContext } from '../llm-usage/llm-usage.context';
import { WebChatAnswer, WebChatCitation, WebChatComplexity } from './web-chat.types';

/**
 * How many past Excel conversations are pulled into an answer's context.
 * Bounded deliberately: a heavy add-in user can have hundreds of threads
 * inside the 90-day window, and stuffing all of them into one prompt would
 * blow the token budget and bury the relevant thread in noise. Newest-first
 * is the right cut because a question about "my work" almost always means
 * recent work; an older thread is still reachable by opening it directly,
 * which scopes the answer to that one conversation (see `conversationId`).
 */
const MAX_CONTEXT_CONVERSATIONS = 8;

/** Messages taken from each conversation — enough for gist, not the full transcript. */
const MAX_MESSAGES_PER_CONVERSATION = 12;

/** Per-message content cap, so one pasted-in giant cell dump can't dominate. */
const MAX_MESSAGE_CHARS = 600;

/**
 * Heuristic complexity split, matching the existing catalog's ask-mode tiers
 * (FORMULA_QA_SIMPLE = 2 credits, FORMULA_QA_COMPLEX = 5). Runs BEFORE the
 * LLM call so the gate check and the user-facing "this will cost N" estimate
 * agree with what is actually debited — deriving price from the model's own
 * output would make the pre-send estimate a guess, which CREDIT_SYSTEM.md
 * CD-1 exists to avoid ("a user should be able to know in advance what an
 * action costs").
 */
const COMPLEX_SIGNALS = [
  'compare',
  'across',
  'all my',
  'every',
  'summar',
  'trend',
  'why did',
  'reconcil',
  'audit',
  'between',
  'difference',
];

@Injectable()
export class WebChatService {
  private readonly logger = new Logger(WebChatService.name);

  constructor(
    @InjectModel(Conversation.name)
    private readonly conversationModel: Model<ConversationDocument>,
    private readonly openRouter: OpenRouterService,
    private readonly usageBilling: UsageBillingService,
  ) {}

  /**
   * Classifies a question as simple or complex. Exported behavior (via
   * `estimateCost`) so the client can show the same number before sending
   * that the server will actually charge.
   */
  classify(question: string, scopedToOneConversation: boolean): WebChatComplexity {
    // A question scoped to a single opened conversation is inherently narrower
    // — one transcript, no cross-thread synthesis — so it prices as simple
    // regardless of wording.
    if (scopedToOneConversation) {
      return 'simple';
    }
    const normalized = question.toLowerCase();
    const hasComplexSignal = COMPLEX_SIGNALS.some((signal) => normalized.includes(signal));
    return hasComplexSignal || question.length > 220 ? 'complex' : 'simple';
  }

  private actionTypeFor(complexity: WebChatComplexity): CreditActionType {
    return complexity === 'complex' ? 'FORMULA_QA_COMPLEX' : 'FORMULA_QA_SIMPLE';
  }

  /**
   * Pre-send estimate for the composer's "~N credits" hint. Billing is now
   * usage-based (TASKS.md #341/#343) — the real charge is the model's actual
   * token cost, not a fixed catalog price — so this reuses the old
   * FORMULA_QA_SIMPLE/COMPLEX prices only as a rough heads-up figure. `ask`
   * does not read this value and can charge a different amount.
   */
  estimateCost(question: string, scopedToOneConversation: boolean): {
    complexity: WebChatComplexity;
    credits: number;
  } {
    const complexity = this.classify(question, scopedToOneConversation);
    return { complexity, credits: resolveCreditCost(this.actionTypeFor(complexity)) };
  }

  /**
   * Answers `question` over the user's stored Excel conversations.
   *
   * Billing mirrors the Excel add-in path (TASKS.md #341): gate on balance >
   * 0 before any LLM spend, then charge the real provider cost of the call
   * once an answer exists. A model/provider failure therefore costs the user
   * nothing — the gate never opens without an answer landing, and a thrown
   * `complete()` never reaches the settle step below.
   */
  async ask(
    userId: string,
    question: string,
    options: { conversationId?: string } = {},
  ): Promise<WebChatAnswer> {
    const scoped = Boolean(options.conversationId);
    const complexity = this.classify(question, scoped);

    if (!(await this.usageBilling.canStart(userId))) {
      throw new InsufficientCreditError(userId, 0);
    }

    const conversations = await this.loadContext(userId, options.conversationId);
    const citations: WebChatCitation[] = conversations.map((conversation) => ({
      conversationId: conversation.conversationId,
      title: conversation.title?.trim() || 'Untitled session',
    }));

    const context: LlmUsageContext = { promptId: randomUUID(), userId };
    const answer = await runWithLlmUsageContext(context, () =>
      this.openRouter.complete({
        systemPrompt: this.buildSystemPrompt(),
        userMessage: this.buildUserMessage(question, conversations),
        // Ask-mode Q&A over an existing transcript is comprehension, not
        // planning or code generation — MEDIUM matches what the add-in's own
        // read-only data-query lane uses for the same class of work.
        tier: 'medium',
        temperature: 0.2,
        maxTokens: complexity === 'complex' ? 1200 : 700,
      }),
    );

    const settlement = await this.usageBilling.settle(context);
    if (settlement && settlement.debited < settlement.requested) {
      this.logger.warn(`Web chat usage debit short for user ${userId} — balance exhausted.`);
    }

    const balances = settlement?.balances;
    const newBalance = balances
      ? balances.planCredits + balances.purchasedCredits + balances.oneTimeCredits
      : 0;

    return {
      answer: answer.trim(),
      citations,
      complexity,
      creditsDeducted: settlement?.debited ?? 0,
      newBalance,
    };
  }

  /**
   * Loads the Excel conversations that form the answer's context — either one
   * specific thread the user has opened, or their most recent threads.
   *
   * `userId` is always part of the filter, including in the single-conversation
   * case: without it, passing someone else's conversationId would read their
   * transcript. This matches the ownership discipline the conversation schema's
   * own docblock describes.
   */
  private async loadContext(
    userId: string,
    conversationId?: string,
  ): Promise<StoredConversation[]> {
    if (conversationId) {
      const one = await this.conversationModel
        .findOne({ conversationId, userId })
        .lean<StoredConversation>();
      return one ? [one] : [];
    }

    return this.conversationModel
      .find({ userId })
      .sort({ updatedAt: -1 })
      .limit(MAX_CONTEXT_CONVERSATIONS)
      .lean<StoredConversation[]>();
  }

  private buildSystemPrompt(): string {
    return [
      'You are Cellix\'s assistant on the web. The user works in Excel through the Cellix add-in,',
      'and you are answering questions ABOUT their past Cellix sessions.',
      '',
      'Rules:',
      '- Answer only from the session transcripts provided below. They are the user\'s own history.',
      '- If the transcripts do not contain the answer, say so plainly and suggest opening the',
      '  relevant sheet in Excel. Never invent a value, formula, or change that is not in them.',
      '- You CANNOT modify any workbook from here. This surface is read-only. If the user asks for',
      '  an edit, explain the change you would make and tell them to run it from the Excel add-in.',
      '- Refer to sessions by their title, not by id.',
      '- Be concise and concrete. Prefer naming the actual sheet, column, or formula involved.',
    ].join('\n');
  }

  private buildUserMessage(question: string, conversations: StoredConversation[]): string {
    if (conversations.length === 0) {
      return [
        'The user has no stored Cellix sessions in the retention window.',
        '',
        `Question: ${question}`,
      ].join('\n');
    }

    const rendered = conversations
      .map((conversation) => this.renderConversation(conversation))
      .join('\n\n---\n\n');

    return [
      "The user's recent Cellix sessions:",
      '',
      rendered,
      '',
      '---',
      '',
      `Question: ${question}`,
    ].join('\n');
  }

  private renderConversation(conversation: StoredConversation): string {
    const lines: string[] = [
      `SESSION: ${conversation.title?.trim() || 'Untitled session'}`,
      `Last active: ${this.formatDate(conversation.updatedAt)}`,
    ];

    // sheetSnapshot is what lets the assistant answer structural questions
    // ("what columns did that sheet have?") without live workbook access,
    // which is unavailable from a browser by construction.
    if (conversation.sheetSnapshot) {
      const { rowCount, columnCount, headers } = conversation.sheetSnapshot;
      lines.push(
        `Sheet: ${rowCount} rows x ${columnCount} columns. Headers: ${headers.join(', ')}`,
      );
    }

    const messages = (conversation.messages ?? []).slice(-MAX_MESSAGES_PER_CONVERSATION);
    for (const message of messages) {
      const speaker = message.role === 'user' ? 'User' : 'Cellix';
      lines.push(`${speaker}: ${this.truncate(message.content)}`);

      // What Cellix actually DID is often the real answer to "what changed?",
      // and it lives in metadata.actions rather than the prose content.
      const actionCount = message.metadata?.actions?.length ?? 0;
      if (actionCount > 0) {
        lines.push(`  (applied ${actionCount} change${actionCount === 1 ? '' : 's'} to the workbook)`);
      }
    }

    return lines.join('\n');
  }

  private truncate(value: string): string {
    const normalized = (value ?? '').replace(/\s+/g, ' ').trim();
    return normalized.length > MAX_MESSAGE_CHARS
      ? `${normalized.slice(0, MAX_MESSAGE_CHARS)}…`
      : normalized;
  }

  private formatDate(value: Date | undefined): string {
    if (!value) return 'unknown';
    return new Date(value).toISOString().slice(0, 10);
  }
}
