import { Injectable } from '@nestjs/common';
import { isFindLookupMessage } from '../utils/find-query-parser.util';
import { IntentClassification, IntentType } from '../types/sheet-actions.types';

@Injectable()
export class IntentClassifierService {
  classify(message: string): IntentClassification {
    const lower = message.toLowerCase().trim();

    if (this.isFixIntent(lower)) {
      return { intent: 'FIX', confidence: 'high', subIntent: this.detectFixType(lower) };
    }

    if (this.isExplainIntent(lower)) {
      return { intent: 'EXPLAIN', confidence: 'high' };
    }

    if (this.isFormulaHelpIntent(lower)) {
      return { intent: 'FORMULA_HELP', confidence: 'high' };
    }

    if (this.isDataQuestionIntent(lower)) {
      return { intent: 'DATA_QUESTION', confidence: 'high', subIntent: this.detectDataQuery(lower) };
    }

    if (this.isActionIntent(lower)) {
      return { intent: 'ACTION', confidence: 'high', subIntent: this.detectActionType(lower) };
    }

    return { intent: 'DATA_QUESTION', confidence: 'low' };
  }

  private isFixIntent(lower: string): boolean {
    return (
      /\b(fix|repair|broken|error|#ref|#value|#n\/a|#div\/0|#name|#num|circular reference|not sorting|sum returns 0|isn't summing|doesn't sum)\b/.test(
        lower,
      ) || lower.includes('something is wrong')
    );
  }

  private detectFixType(lower: string): string {
    if (lower.includes('#ref')) return 'ref_error';
    if (lower.includes('#n/a') || lower.includes('vlookup')) return 'na_error';
    if (lower.includes('#value')) return 'value_error';
    if (lower.includes('#div')) return 'div_error';
    if (lower.includes('circular')) return 'circular_ref';
    if (lower.includes('not sorting') || lower.includes("aren't sorting")) return 'text_dates';
    if (lower.includes('sum returns 0') || lower.includes("aren't summing")) return 'text_numbers';
    return 'generic';
  }

  private isExplainIntent(lower: string): boolean {
    return (
      /\b(what does|describe|tell me about|explain|what is in|what are all|what named ranges|what filter|what sheets|what is the data type|are there any merged|overview|summarize)\b/.test(
        lower,
      ) || /\bwhat('s| is) (in|on) (this|the) (sheet|spreadsheet|workbook)\b/.test(lower)
    );
  }

  private isFormulaHelpIntent(lower: string): boolean {
    return (
      /\b(write|create|generate|give me|build)\b.*\b(formula|function)\b/.test(lower) ||
      /\b(formula|function)\b.*\b(for|to|that)\b/.test(lower) ||
      /\bcalculate\b.*\b(gst|igst|cgst|sgst|tds|percentage|percent)\b/.test(lower)
    );
  }

  isFindLookupIntent(lower: string): boolean {
    return isFindLookupMessage(lower);
  }

  private isDataQuestionIntent(lower: string): boolean {
    return (
      this.isFindLookupIntent(lower) ||
      /\b(how many|what is the total|what is the average|what is the highest|what is the lowest|what is the maximum|what is the minimum|which rows|which supplier|which column|are there duplicate|what percentage|count of|number of)\b/.test(
        lower,
      ) ||
      (/\b(total|sum|average|max|min|count|highest|lowest|maximum|minimum)\b/.test(lower) &&
        !this.isWriteIntent(lower))
    );
  }

  private detectDataQuery(lower: string): string {
    if (this.isFindLookupIntent(lower)) return 'find';
    if (/\b(average|mean)\b/.test(lower)) return 'average';
    if (/\b(max|maximum|highest|largest)\b/.test(lower)) return 'max';
    if (/\b(min|minimum|lowest|smallest)\b/.test(lower)) return 'min';
    if (/\b(count|how many)\b/.test(lower)) return 'count';
    if (/\b(blank|empty)\b/.test(lower)) return 'blank';
    if (/\b(duplicate)\b/.test(lower)) return 'duplicate';
    if (/\b(percentage|percent)\b/.test(lower)) return 'percentage';
    return 'sum';
  }

  private isActionIntent(lower: string): boolean {
    return this.isWriteIntent(lower);
  }

  private isWriteIntent(lower: string): boolean {
    return (
      /\b(add|create|insert|delete|remove|rename|copy|move|hide|unhide|show|format|apply|make|set|put|write|change|update|merge|unmerge|clear|sort|filter|freeze|paste|fill|highlight|lock|unlock|protect|resize|wrap|generate|populate|build|fill in|dummy|sample|random|seed|mock)\b/.test(
        lower,
      ) || /\bdo (this|it|that)\b/.test(lower)
    );
  }

  private detectActionType(lower: string): string {
    if (/\b(sheet|tab)\b/.test(lower)) return 'sheet';
    if (/\b(row|rows)\b/.test(lower)) return 'row';
    if (/\b(column|col)\b/.test(lower)) return 'column';
    if (/\b(format|bold|italic|colour|color|currency|percentage|border|align|wrap)\b/.test(lower))
      return 'format';
    if (/\b(formula|sum|average|if|vlookup|countif|sumif)\b/.test(lower)) return 'formula';
    if (/\b(merge|unmerge)\b/.test(lower)) return 'cell';
    if (/\b(sort|filter|find|replace)\b/.test(lower)) return 'data';
    return 'general';
  }
}

export function intentRequiresApproval(intent: IntentType): boolean {
  return intent === 'ACTION' || intent === 'FIX' || intent === 'FORMULA_HELP';
}

export function intentIsReadOnly(intent: IntentType): boolean {
  return intent === 'EXPLAIN' || intent === 'DATA_QUESTION';
}
