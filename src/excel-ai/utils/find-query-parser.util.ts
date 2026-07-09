const FIND_LOOKUP_PATTERN =
  /\b(find|search|locate|look up|lookup|show me|where is|get me|fetch|pull up|bring up|list rows|list all rows|show rows|show all rows)\b/i;

const FIND_VERB_CAPTURE =
  /^(?:find|search|locate|look\s*up|lookup|show\s+me|where\s+is|get\s+me|fetch|list(?:\s+all)?)\s+/i;

/** Clauses that turn a find into a write/export follow-up (handled separately from read-only find). */
const FIND_FOLLOW_ON_SPLIT =
  /\s+and\s+(?=(?:create|add|make|build|copy|move|put|paste|export|transfer|generate)\b)/i;

const ROW_EXPORT_PATTERN =
  /\b(?:create|add|make|build)\s+(?:a\s+)?(?:new\s+)?sheet\b|\bcopy\b[^.]{0,48}\brows?\b|\b(?:put|paste|move|export|transfer)\b[^.]{0,48}\b(?:rows?|data|values?|records?)\b|\b(?:to|into)\s+(?:a\s+)?(?:new\s+)?sheet\b/i;

const MULTI_VALUE_FIND_PATTERN =
  /\b(?:values?|entries|items|amounts|numbers|records)\b|(?:,|\band\b).*?\d/;

export type LocalFindRoute = 'none' | 'read_only' | 'export_rows';

export function isFindLookupMessage(message: string): boolean {
  return FIND_LOOKUP_PATTERN.test(message.trim());
}

export function wantsRowExport(message: string): boolean {
  return ROW_EXPORT_PATTERN.test(message.trim());
}

export function resolveLocalFindRoute(message: string): LocalFindRoute {
  const trimmed = message.trim();
  if (!isFindLookupMessage(trimmed)) return 'none';
  if (wantsRowExport(trimmed)) return 'export_rows';
  return 'read_only';
}

export function stripFindFollowOnClauses(message: string): string {
  const parts = message.split(FIND_FOLLOW_ON_SPLIT);
  return parts[0]?.trim() ?? message.trim();
}

function extractNumericTokens(message: string): string[] {
  const tokens = message.match(/\d[\d,]*(?:\.\d+)?/g) ?? [];
  return [
    ...new Set(
      tokens
        .map((token) => token.replace(/,/g, ''))
        .filter(Boolean),
    ),
  ];
}

function isMultiValueFindQuery(message: string, numericTokens: string[]): boolean {
  if (numericTokens.length < 2) return false;
  const lower = message.toLowerCase();
  if (MULTI_VALUE_FIND_PATTERN.test(lower)) return true;
  return /,\s*\d/.test(message) || /\band\s+\d/.test(lower);
}

function extractTextFindPhrase(message: string): string | null {
  const withoutVerb = message.replace(FIND_VERB_CAPTURE, '').trim();
  if (!withoutVerb) return null;

  const match = withoutVerb.match(
    /^(?:all\s+)?(?:the\s+)?(?:rows?\s+)?(?:with|where|containing|matching|for|having)?\s*(.+?)\s*$/i,
  );
  if (!match?.[1]) return null;

  let phrase = match[1].trim();
  phrase = phrase.replace(/\s+(?:rows?|records?|entries?)\s*$/i, '').trim();
  return phrase || null;
}

/** Aggregation prompts (total/sum/average) are not cell lookup searches. */
export function isDataAggregationMessage(message: string): boolean {
  const lower = message.toLowerCase().trim();
  if (isFindLookupMessage(lower)) return false;
  return (
    /\b(what is the total|what's the total|what is the sum|what's the sum|how much|how many|what is the average|what's the average|what is the highest|what is the lowest|what is the maximum|what is the minimum)\b/.test(
      lower,
    ) ||
    (/\b(total|sum|average|count|maximum|minimum|highest|lowest)\b/.test(lower) &&
      /\b(cgst|sgst|igst|gst|amount|invoice|column|row)\b/.test(lower))
  );
}

/** Parse one or more search terms from a natural-language find prompt. */
export function parseFindSearchTerms(message: string): string[] {
  if (isDataAggregationMessage(message)) return [];

  const quoted = /"([^"]+)"/.exec(message);
  if (quoted?.[1]?.trim()) return [quoted[1].trim()];

  const scoped = stripFindFollowOnClauses(message);

  const numericTokens = extractNumericTokens(scoped);
  if (numericTokens.length > 0) {
    if (isMultiValueFindQuery(scoped, numericTokens)) {
      return numericTokens;
    }
    return [numericTokens[numericTokens.length - 1]!];
  }

  const textPhrase = extractTextFindPhrase(scoped);
  return textPhrase ? [textPhrase] : [];
}

export function sanitizeSheetName(name: string): string {
  const cleaned = name
    .replace(/[[\]:*?/\\]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 31);
  return cleaned || 'Filtered rows';
}

export function suggestExportSheetName(message: string, searchLabel: string): string {
  const quotedName = /\b(?:sheet\s+)?(?:named|called)\s+["']([^"']+)["']/i.exec(message);
  if (quotedName?.[1]?.trim()) {
    return sanitizeSheetName(quotedName[1].trim());
  }

  const explicitName =
    /\b(?:sheet\s+)?(?:named|called)\s+([A-Za-z][A-Za-z0-9 _-]{0,28})(?:\s*$|\s*\.|\s+with\b)/i.exec(
      message,
    );
  if (explicitName?.[1]?.trim()) {
    return sanitizeSheetName(explicitName[1].trim());
  }

  return sanitizeSheetName(searchLabel);
}
