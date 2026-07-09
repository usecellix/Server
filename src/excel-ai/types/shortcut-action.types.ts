import { SheetActionPayload } from './sheet-actions.types';

export type ShortcutHandler = (message: string, activeSheetName?: string) => SheetActionPayload[] | null;

export interface ShortcutActionDefinition {
  id: string;
  description: string;
  handler: ShortcutHandler;
}
