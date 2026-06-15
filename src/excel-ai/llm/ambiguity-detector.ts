import { ClarificationPayload, UserIntent, WorkbookContext } from '../../types/cellix.types';

export interface AmbiguityResult {
  score: number;
  needsClarification: boolean;
  question?: string;
  suggestions?: string[];
  reason?: string;
}

export interface AmbiguityDetectionOutcome {
  clarification: ClarificationPayload | null;
  score: number;
  lowConfidence: boolean;
  reasons: string[];
}

const VAGUE_VERBS = ['update', 'change', 'fix', 'modify', 'edit', 'adjust', 'do'];
const VAGUE_REFS = ['it', 'that', 'this', 'those', 'here', 'there', 'them'];
const VAGUE_SCOPE = ['some', 'few', 'many', 'several', 'most', 'all', 'rest'];

interface RuleScore {
  score: number;
  reason: string;
}

const INTENT_PATTERNS: [UserIntent, RegExp][] = [
  ['create_data', /\b(create|generate|make|build|add|insert|fill)\b.*\b(data|rows?|sheet|table)\b/i],
  ['modify_data', /\b(update|change|edit|modify|set|replace)\b/i],
  ['format', /\b(bold|italic|color|font|border|align|format|style)\b/i],
  ['formula', /\b(formula|function|sum|average|count|if|vlookup|xlookup)\b/i],
  ['sort_filter', /\b(sort|filter|order|arrange)\b/i],
  ['analyze', /\b(compare|analyze|analysis|difference|summary|total)\b/i],
  ['delete', /\b(delete|remove|clear|erase|drop)\b/i],
];

export function classifyIntent(prompt: string): UserIntent {
  for (const [intent, pattern] of INTENT_PATTERNS) {
    if (pattern.test(prompt)) return intent;
  }
  return 'other';
}

function ruleBasedScore(
  prompt: string,
  context: WorkbookContext,
  hasHistory: boolean,
): RuleScore[] {
  const lower = prompt.toLowerCase().trim();
  const tokens = lower.split(/\s+/);
  const scores: RuleScore[] = [];
  const primarySheet = context.sheets[0];

  if (tokens.length < 4) {
    scores.push({ score: 50, reason: 'Prompt is very short' });
  }

  if (VAGUE_VERBS.some((v) => tokens[0] === v) && tokens.length < 8) {
    scores.push({ score: 40, reason: 'Prompt starts with a vague verb' });
  }

  if (!hasHistory && VAGUE_REFS.some((r) => tokens.includes(r))) {
    scores.push({ score: 45, reason: 'Uses vague pronoun without context' });
  }

  if (VAGUE_SCOPE.some((s) => tokens.includes(s))) {
    scores.push({ score: 30, reason: 'Vague quantity word' });
  }

  const hasMultipleColumns = (primarySheet?.colCount ?? 0) > 3;
  const hasColumnRef =
    primarySheet?.headers?.some((h) => h && lower.includes(h.toLowerCase())) ?? false;
  const hasCellRef = /[A-Z]+\d+/.test(prompt.toUpperCase());

  if (hasMultipleColumns && !hasColumnRef && !hasCellRef) {
    if (/sort|filter|group|sum|average|max|min/.test(lower)) {
      scores.push({ score: 55, reason: 'Aggregation without column reference' });
    }
  }

  const hasMultipleSheets = context.sheets.length > 1;
  const sheetNamesInPrompt = context.sheets.some((s) =>
    lower.includes(s.sheetName.toLowerCase()),
  );
  if (hasMultipleSheets && !sheetNamesInPrompt) {
    scores.push({ score: 25, reason: 'Multiple sheets but no sheet named' });
  }

  if (/^(create|add|insert)/.test(lower) && tokens.length < 6) {
    scores.push({ score: 45, reason: 'Create/add without specifics' });
  }

  return scores;
}

const FIND_LOOKUP_PATTERN =
  /\b(find|search|locate|look up|lookup|show me|where is|get me|fetch|pull up|bring up|list rows|list all rows|show rows|show all rows)\b/i;

export function scoreAmbiguity(
  prompt: string,
  context: WorkbookContext,
  hasHistory: boolean,
): { score: number; reasons: string[] } {
  // Find/search/locate queries are self-contained — never ask for clarification.
  if (FIND_LOOKUP_PATTERN.test(prompt)) {
    return { score: 0, reasons: [] };
  }

  const rules = ruleBasedScore(prompt, context, hasHistory);
  if (rules.length === 0) return { score: 0, reasons: [] };
  const avg = rules.reduce((sum, rule) => sum + rule.score, 0) / rules.length;
  return {
    score: Math.min(100, Math.round(avg)),
    reasons: rules.map((rule) => rule.reason),
  };
}

function buildFallbackClarification(
  prompt: string,
  context: WorkbookContext,
): { question: string; suggestions: string[] } {
  const lower = prompt.toLowerCase();
  const headers = context.sheets[0]?.headers?.filter((h) => h.trim()) ?? [];

  if (/sort|order|arrange/.test(lower) && headers.length > 0) {
    return {
      question: 'Which column should I sort by?',
      suggestions: headers.slice(0, 4),
    };
  }

  if (/filter/.test(lower) && headers.length > 0) {
    return {
      question: 'Which column should I filter on?',
      suggestions: headers.slice(0, 4),
    };
  }

  if (/sum|average|total|max|min/.test(lower) && headers.length > 0) {
    return {
      question: 'Which column should I calculate on?',
      suggestions: headers.slice(0, 4),
    };
  }

  if (context.sheets.length > 1) {
    return {
      question: 'Which sheet should I use?',
      suggestions: context.sheets.map((sheet) => sheet.sheetName).slice(0, 4),
    };
  }

  return {
    question: 'Could you give more detail about what you want to do?',
    suggestions: headers.slice(0, 3),
  };
}

export async function generateClarificationQuestion(
  prompt: string,
  context: WorkbookContext,
  quickCall: (systemPrompt: string, userMessage: string) => Promise<string>,
): Promise<{ question: string; suggestions: string[] }> {
  const sheetSummary = context.sheets
    .map(
      (sheet) =>
        `Sheet "${sheet.sheetName}": ${sheet.headers.join(', ')} (${sheet.rowCount} rows)`,
    )
    .join('\n');

  const systemPrompt = `You are a clarification assistant for an Excel AI tool.
The user gave a prompt that is slightly ambiguous. Your job is to ask ONE specific, short clarifying question.
Also provide 2–4 short answer suggestions that cover the most likely options.

Workbook context:
${sheetSummary}

Respond ONLY in this JSON format (no markdown, no explanation):
{"question":"...","suggestions":["...","...","..."]}`;

  try {
    const raw = await quickCall(systemPrompt, prompt);
    const parsed = JSON.parse(raw.trim()) as {
      question?: string;
      suggestions?: string[];
    };
    return {
      question: parsed.question ?? 'Could you be more specific?',
      suggestions: Array.isArray(parsed.suggestions) ? parsed.suggestions : [],
    };
  } catch {
    return buildFallbackClarification(prompt, context);
  }
}

export async function detectAmbiguity(
  prompt: string,
  context: WorkbookContext,
  conversationHistory: Array<{ role: string; content: string }>,
  quickCall?: (systemPrompt: string, userMessage: string) => Promise<string>,
): Promise<AmbiguityDetectionOutcome> {
  const hasHistory = conversationHistory.length > 0;
  const { score, reasons } = scoreAmbiguity(prompt, context, hasHistory);

  if (score < 40) {
    return { clarification: null, score, lowConfidence: false, reasons };
  }

  let question = 'Could you give more detail about what you want to do?';
  let suggestions: string[] = [];

  if (score >= 40 && score < 66 && quickCall) {
    const result = await generateClarificationQuestion(prompt, context, quickCall);
    question = result.question;
    suggestions = result.suggestions;
  } else if (score >= 66) {
    if (quickCall) {
      const result = await generateClarificationQuestion(prompt, context, quickCall);
      question = result.question;
      suggestions = result.suggestions;
    } else {
      const fallback = buildFallbackClarification(prompt, context);
      question = fallback.question;
      suggestions = fallback.suggestions;
    }
  }

  if (score < 66) {
    return { clarification: null, score, lowConfidence: true, reasons };
  }

  return {
    clarification: {
      question,
      suggestions,
      ambiguityScore: score,
    },
    score,
    lowConfidence: false,
    reasons,
  };
}
