import { Action } from '../../agents/types/agent.types';
import {
  DomainException,
  SourceRef,
} from '../../domain-tools/types/domain-tool.types';
import { extractFormulaPrecedents } from './formula-precedents.util';

export interface ProvenanceContext {
  /** Workbook id / name used as SourceRef.documentId for workbook refs */
  workbookId?: string;
  activeSheetName?: string;
  sourceRefs?: SourceRef[];
  exceptionFlags?: DomainException[];
  confidence?: number;
  /** When true, createPreview requires non-empty sourceRefs */
  fromDomainTool?: boolean;
}

export function buildWorkbookSourceRefsFromActions(
  actions: Action[],
  workbookId: string,
  activeSheetName?: string,
): SourceRef[] {
  const seen = new Set<string>();
  const refs: SourceRef[] = [];

  for (const action of actions) {
    const formula =
      action.type === 'SET_FORMULA'
        ? action.formula
        : action.type === 'SET_CELL' && typeof action.value === 'string' && action.value.startsWith('=')
          ? action.value
          : undefined;
    if (!formula) continue;

    const sheet = action.sheetName || activeSheetName;
    for (const precedent of extractFormulaPrecedents(formula, sheet)) {
      if (seen.has(precedent)) continue;
      seen.add(precedent);
      refs.push({
        documentType: 'workbook',
        documentId: workbookId,
        rowOrLine: precedent,
      });
    }
  }

  return refs;
}

export function assertDomainToolProvenance(provenance?: ProvenanceContext): void {
  if (!provenance?.fromDomainTool) return;
  if (!provenance.sourceRefs?.length) {
    throw new Error(
      'Domain-tool-backed writes must include non-empty sourceRefs before createPreview()',
    );
  }
}
