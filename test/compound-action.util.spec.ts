import {
  buildCompoundCreateAndSortActions,
  buildCompoundFallbackSubtasks,
  buildDeterministicSubtaskActions,
  detectCreateNewSheet,
  detectSortIntent,
} from '../src/agents/utils/compound-action.util';
import { buildSortFallbackAction } from '../src/agents/utils/sort-action.util';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

const gstContext: WorkbookContext = {
  activeSheetName: 'Invoices',
  sheets: [
    {
      name: 'Invoices',
      usedRange: 'A1:M339',
      rowCount: 339,
      columnCount: 13,
      values: [
        [
          'Invoice No',
          'Date',
          'Customer',
          'Amount',
          'SGST',
          'CGST',
          'IGST',
          'Total',
          'Status',
          'Region',
          'Type',
          'Notes',
          'Ref',
        ],
        ...Array.from({ length: 338 }, (_, index) => [
          `INV-${index + 1}`,
          '2024-01-01',
          'Customer',
          1000 + index,
          90,
          100 + (index % 50),
          0,
          1190 + index,
          'Paid',
          'South',
          'B2B',
          '',
          `R-${index}`,
        ]),
      ],
      formulas: [],
      numberFormats: [],
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('sort-action.util CGST phrasing', () => {
  it('extracts CGST from "sort the values of CGST in ascending order"', () => {
    const subtask: SubTask = {
      id: 's1',
      description: 'sort the values of CGST in ascending order',
      targetSheet: 'Invoices',
      dependsOn: [],
      estimatedActions: 1,
    };

    const action = buildSortFallbackAction(subtask, gstContext);
    expect(action).not.toBeNull();
    expect(action?.columnName).toBe('CGST');
    expect(action?.ascending).toBe(true);
    expect(action?.key).toBe(5);
  });
});

describe('compound-action.util', () => {
  const prompt = 'create new sheet and sort the values of CGST in ascending order';

  it('detects create sheet and sort intents', () => {
    expect(detectCreateNewSheet(prompt)).toBe(true);
    expect(detectSortIntent(prompt)).toBe(true);
  });

  it('builds split fallback subtasks', () => {
    const subtasks = buildCompoundFallbackSubtasks(prompt, gstContext);
    expect(subtasks).toHaveLength(2);
    expect(subtasks?.[0].description).toContain('Create new sheet');
    expect(subtasks?.[1].dependsOn).toEqual(['s1']);
    expect(subtasks?.[1].description).toContain('CGST');
  });

  it('builds compound create + sort actions for a single subtask', () => {
    const subtask: SubTask = {
      id: 's1',
      description: prompt,
      targetSheet: 'Invoices',
      dependsOn: [],
      estimatedActions: 2,
    };

    const result = buildCompoundCreateAndSortActions(prompt, gstContext, subtask);
    expect(result?.actions).toHaveLength(2);
    expect(result?.actions[0].type).toBe('ADD_SHEET');
    expect(result?.actions[1].type).toBe('SORT_RANGE');
    expect(result?.actions[1].columnName).toBe('CGST');
    expect(result?.isDone).toBe(true);
  });

  it('builds deterministic sort-only actions for split subtask s2', () => {
    const newSheetName = 'CGST Sorted';
    const sortContext: WorkbookContext = {
      ...gstContext,
      sheets: [
        ...gstContext.sheets,
        { ...gstContext.sheets[0], name: newSheetName },
      ],
    };
    const subtask: SubTask = {
      id: 's2',
      description: 'sort the values of CGST in ascending order on sheet "CGST Sorted"',
      targetSheet: newSheetName,
      dependsOn: ['s1'],
      estimatedActions: 1,
    };

    const result = buildDeterministicSubtaskActions(subtask, sortContext);
    expect(result?.actions).toHaveLength(1);
    expect(result?.actions[0].type).toBe('SORT_RANGE');
    expect(result?.actions[0].sheetName).toBe(newSheetName);
  });

  it('builds deterministic create-sheet action for split subtask s1', () => {
    const subtask: SubTask = {
      id: 's1',
      description: 'Create new sheet "CGST Sorted" as a copy of "Invoices"',
      targetSheet: 'Invoices',
      dependsOn: [],
      estimatedActions: 1,
    };

    const result = buildDeterministicSubtaskActions(subtask, gstContext);
    expect(result?.actions).toHaveLength(1);
    expect(result?.actions[0].type).toBe('ADD_SHEET');
    expect(result?.actions[0].name).toBe('CGST Sorted');
    expect(result?.actions[0].copyFrom).toBe('Invoices');
  });

  it('does not treat sheet name "Sorted" as a sort instruction', () => {
    expect(
      detectSortIntent('Create new sheet "CGST Sorted" as a copy of "Invoices"'),
    ).toBe(false);
  });

  it('defers create sheet with dummy data to the LLM (no deterministic shortcut)', () => {
    const subtask: SubTask = {
      id: 's1',
      description: 'Create a new sheet named Summary with gst dummy data',
      targetSheet: 'Invoices',
      dependsOn: [],
      estimatedActions: 2,
    };

    expect(buildDeterministicSubtaskActions(subtask, gstContext)).toBeNull();
  });

  it('builds a blank sheet (no copy) for "create an empty sheet named Cellix"', () => {
    const subtask: SubTask = {
      id: 's1',
      description: 'Create an empty sheet named Cellix',
      targetSheet: 'Invoices',
      dependsOn: [],
      estimatedActions: 1,
    };

    const result = buildDeterministicSubtaskActions(subtask, gstContext);
    expect(result?.actions).toHaveLength(1);
    expect(result?.actions[0].type).toBe('ADD_SHEET');
    expect(result?.actions[0].name).toBe('Cellix');
    expect(result?.actions[0].copyFrom).toBeUndefined();
  });
});
