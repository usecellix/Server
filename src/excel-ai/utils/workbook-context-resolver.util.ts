import { WorkbookContext, SheetSnapshot } from '../../types/cellix.types';
import { ConversationRequestDto, WorkbookContextDto } from '../dto/conversation-request.dto';
import { WorkbookContextInput } from '../services/conversation-engine.service';
import { SheetAnalysis } from '../services/sheet-analyzer.service';

function isSheetSnapshot(value: unknown): value is SheetSnapshot {
  return (
    typeof value === 'object' &&
    value !== null &&
    'sheetName' in value &&
    'headers' in value &&
    Array.isArray((value as SheetSnapshot).headers)
  );
}

function isRichWorkbookContext(value: unknown): value is WorkbookContext {
  return (
    typeof value === 'object' &&
    value !== null &&
    'sheets' in value &&
    Array.isArray((value as WorkbookContext).sheets) &&
    (value as WorkbookContext).sheets.length > 0 &&
    isSheetSnapshot((value as WorkbookContext).sheets[0])
  );
}

export function resolveWorkbookContext(
  request: ConversationRequestDto,
  analysis: SheetAnalysis,
  sheetData: unknown[][],
): WorkbookContext {
  if (isRichWorkbookContext(request.workbookContext)) {
    const rich = request.workbookContext;
    const promptContext = request.promptContext ?? rich.prompt_context;
    if (promptContext && promptContext !== rich.prompt_context) {
      return { ...rich, prompt_context: promptContext };
    }
    return rich;
  }

  const legacy = request.workbookContext as
    | { activeSheet?: string; sheets?: string[] }
    | undefined;
  const activeSheet = legacy?.activeSheet ?? 'Sheet1';
  const sheetNames = legacy?.sheets?.length ? legacy.sheets : [activeSheet];

  const sampleData = sheetData.slice(1, 11).map((row) =>
    Array.isArray(row)
      ? row.map((cell) =>
          cell === '' || cell == null ? null : (cell as string | number),
        )
      : [],
  );

  const snapshot: SheetSnapshot = {
    sheetName: activeSheet,
    usedRange: analysis.isEmpty
      ? 'A1'
      : `A1:${analysis.columnLetters[analysis.columnCount - 1] ?? 'A'}${analysis.rowCount}`,
    rowCount: analysis.rowCount,
    colCount: analysis.columnCount,
    headers: analysis.headers,
    sampleData,
    columnMeta: analysis.headers.map((header, index) => {
      const sampleValues = sampleData
        .map((row) => row[index] ?? null)
        .filter((v) => v != null)
        .slice(0, 5);
      const allNumeric = sampleValues.length > 0 && sampleValues.every((v) => typeof v === 'number');
      return {
        index,
        header,
        sampleValues,
        detectedType: allNumeric ? 'number' : sampleValues.length > 0 ? 'text' : 'unknown',
      };
    }),
  };

  return {
    activeSheet,
    sheets: sheetNames.map((name) =>
      name === activeSheet
        ? snapshot
        : {
            ...snapshot,
            sheetName: name,
          },
    ),
    ...(request.promptContext ? { prompt_context: request.promptContext } : {}),
  };
}

export function resolveEngineWorkbookMeta(
  request: ConversationRequestDto,
): WorkbookContextInput | undefined {
  const workbookContext = request.workbookContext;
  if (!workbookContext) {
    return undefined;
  }

  if (isRichWorkbookContext(workbookContext)) {
    return {
      activeSheet: workbookContext.activeSheet,
      sheets: workbookContext.sheets.map((sheet) => sheet.sheetName),
    };
  }

  const legacy = workbookContext as WorkbookContextDto;
  return {
    activeSheet: legacy.activeSheet,
    sheets: legacy.sheets,
  };
}

export function resolveConversationHistory(
  request: ConversationRequestDto,
): Array<{ role: string; content: string }> {
  if (request.conversationHistory?.length) {
    return request.conversationHistory.map((entry) => ({
      role: entry.role,
      content: entry.content,
    }));
  }

  return (request.context?.previousMessages ?? []).map((entry) => ({
    role: entry.role,
    content: entry.content,
  }));
}
