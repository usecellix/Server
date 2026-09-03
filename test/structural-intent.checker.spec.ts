import { StructuralIntentChecker } from '../src/agents/checkers/structural-intent.checker';
import { SubTask, WorkbookContext } from '../src/agents/types/agent.types';

describe('StructuralIntentChecker', () => {
  const checker = new StructuralIntentChecker();

  const emptyWorkbook: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:A1',
        rowCount: 1,
        columnCount: 1,
        values: [['']],
        formulas: [['']],
        numberFormats: [['General']],
        structure: 'unknown',
        headerRowIndex: 0,
      },
    ],
    namedRanges: [],
    tables: [],
  };

  const createMainSubtask: SubTask = {
    id: 's1',
    description: "Create sheet 'Main' if it doesn't exist",
    targetSheet: 'Main',
    dependsOn: [],
    estimatedActions: 1,
  };

  it('fails when CREATE_SHEET lands under a different name than the subtask targets (COMPETITIVE_STUDY_SHORTCUT.md:71 — Main -> Main 2)', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main 2' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(false);
    expect(result.subtaskResults[0].passed).toBe(false);
    expect(result.subtaskResults[0].feedback).toContain('Main');
    expect(result.subtaskResults[0].feedback).toContain('Main 2');
  });

  it('passes when CREATE_SHEET uses the exact requested name', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });

  it('passes when ADD_SHEET (the normalized type) uses the "name" field matching the target', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'ADD_SHEET', name: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });

  it('skips the check entirely when the target sheet already exists (idempotent reuse is correct behavior)', () => {
    const contextWithMain: WorkbookContext = {
      ...emptyWorkbook,
      sheets: [...emptyWorkbook.sheets, { ...emptyWorkbook.sheets[0], name: 'Main' }],
    };

    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          // Executor correctly emitted nothing further, or referenced the existing sheet.
          actions: [],
        },
      ],
      contextWithMain,
    );

    expect(result.passed).toBe(true);
  });

  it('skips subtasks that are not about creating a sheet', () => {
    const result = checker.check(
      [
        {
          subtask: {
            id: 's2',
            description: 'Set Remarks to Cleared where Payment Status = Paid',
            targetSheet: 'Purchase Register',
            dependsOn: [],
            estimatedActions: 1,
          },
          actions: [{ type: 'SET_MATCHING_ROWS', targetColumn: 'Remarks' } as never],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
    expect(result.subtaskResults[0].feedback).toContain('skipped');
  });

  it('ignores CREATE_SHEET actions belonging to other subtasks in the same batch', () => {
    const result = checker.check(
      [
        {
          subtask: createMainSubtask,
          actions: [{ type: 'CREATE_SHEET', sheetName: 'Main' }],
        },
      ],
      emptyWorkbook,
    );

    expect(result.passed).toBe(true);
  });
});
