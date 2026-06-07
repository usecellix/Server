import { SheetAnalysis } from '../services/sheet-analyzer.service';

export function buildStatusMessage(message: string, analysis: SheetAnalysis): string {
  const lower = message.toLowerCase();

  if (analysis.isEmpty) {
    if (/\b(create|generate|populate|add|fill|make|build|dummy|sample)\b/.test(lower)) {
      return 'Creating your table…';
    }
    if (/\b(explain|describe|what)\b/.test(lower)) {
      return 'Checking your sheet…';
    }
    return 'Working on your request…';
  }

  if (/\b(create|generate|populate|add row|insert)\b/.test(lower)) {
    return 'Updating your sheet…';
  }
  if (/\b(total|sum|average|count|how many)\b/.test(lower)) {
    return 'Calculating…';
  }
  if (/\b(format|bold|colour|color|currency)\b/.test(lower)) {
    return 'Applying formatting…';
  }

  return 'Working on your request…';
}
