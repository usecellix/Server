import { SheetContext } from '../agents/types/agent.types';
import { FormulaInsights } from './formula.types';

export type EnrichedSheetContext = SheetContext & {
  formulaInsights?: FormulaInsights;
};

export function buildEnrichedPromptContext(
  basePromptContext: string | undefined,
  sheets: EnrichedSheetContext[],
): string {
  const formulaSections = sheets
    .filter((s) => s.formulaInsights && s.formulaInsights.totalFormulas > 0)
    .map((s) => s.formulaInsights!.llmSummary)
    .join('\n\n');

  const base = basePromptContext?.trim() ?? '';
  if (!formulaSections) return base;
  return base ? `${base}\n\n${formulaSections}` : formulaSections;
}
