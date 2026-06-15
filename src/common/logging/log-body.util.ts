const MAX_LOG_LENGTH = 4096;

type JsonRecord = Record<string, unknown>;

function truncate(value: string): string {
  if (value.length <= MAX_LOG_LENGTH) {
    return value;
  }
  return `${value.slice(0, MAX_LOG_LENGTH)}… [truncated ${value.length - MAX_LOG_LENGTH} chars]`;
}

function isRecord(value: unknown): value is JsonRecord {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function summarizeSheetData(sheetData: unknown): JsonRecord | string {
  if (!Array.isArray(sheetData)) {
    return '[invalid sheetData]';
  }

  const firstRow = sheetData[0];
  const columnCount = Array.isArray(firstRow) ? firstRow.length : 0;
  const headers = Array.isArray(firstRow) ? firstRow.slice(0, 12) : [];

  return {
    rows: sheetData.length,
    columns: columnCount,
    headers,
  };
}

export function sanitizeLogBody(body: unknown): unknown {
  if (body === undefined || body === null) {
    return undefined;
  }

  if (!isRecord(body)) {
    if (typeof body === 'string') {
      return truncate(body);
    }
    return body;
  }

  const sanitized: JsonRecord = { ...body };

  if ('sheetData' in sanitized) {
    sanitized.sheetData = summarizeSheetData(sanitized.sheetData);
  }

  if (isRecord(sanitized.context) && Array.isArray(sanitized.context.previousMessages)) {
    sanitized.context = {
      ...sanitized.context,
      previousMessages: `${sanitized.context.previousMessages.length} message(s)`,
    };
  }

  return sanitized;
}

export function serializeLogBody(body: unknown): string {
  const sanitized = sanitizeLogBody(body);
  if (sanitized === undefined) {
    return '-';
  }
  try {
    return truncate(JSON.stringify(sanitized));
  } catch {
    return '[unserializable]';
  }
}
