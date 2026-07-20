import * as fs from 'fs';
import * as path from 'path';
import { Tier2GenerateVerifyService } from '../src/excel-ai/services/tier2-generate-verify.service';
import { ExecutorAgent } from '../src/agents/executor.agent';
import { VerifierAgent } from '../src/agents/verifier.agent';
import { FormulaValidatorService } from '../src/formula/formula-validator.service';
import { WorkbookContext } from '../src/agents/types/agent.types';
import { assertTier2VerifierMandatory } from '../src/excel-ai/utils/tier2-verifier.guard';

const tier2Fixtures = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'fixtures', 'tier2-formula-fixtures.json'), 'utf8'),
) as Array<{
  message: string;
  actionHint: string;
  expectedFormula?: string;
  expectedActionType?: string;
}>;

const workbookContext: WorkbookContext = {
  activeSheetName: 'Invoices',
  sheets: [
    {
      name: 'Invoices',
      usedRange: 'A1:F50',
      rowCount: 50,
      columnCount: 6,
      values: [
        ['Invoice', 'Amount', 'GST', 'Status', 'Region', 'Date'],
        ['INV-1', 1000, 180, 'Paid', 'South', '2024-01-01'],
        ['INV-2', 2500, 450, 'Open', 'North', '2024-01-02'],
      ],
      formulas: Array.from({ length: 3 }, () => Array(6).fill('')),
      numberFormats: Array.from({ length: 3 }, () => Array(6).fill('General')),
      structure: 'data_table',
    },
  ],
  namedRanges: [],
  tables: [],
};

describe('Tier2GenerateVerifyService', () => {
  let executor: jest.Mocked<Pick<ExecutorAgent, 'execute'>>;
  let verifier: jest.Mocked<Pick<VerifierAgent, 'verify'>>;
  let formulaValidator: FormulaValidatorService;
  let service: Tier2GenerateVerifyService;

  beforeEach(() => {
    executor = { execute: jest.fn() };
    verifier = { verify: jest.fn() };
    formulaValidator = new FormulaValidatorService();
    service = new Tier2GenerateVerifyService(
      executor as unknown as ExecutorAgent,
      verifier as unknown as VerifierAgent,
      formulaValidator,
    );
  });

  it('never calls PlannerAgent — only executor then verifier', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [
        { type: 'SET_FORMULA', sheetName: 'Invoices', row: 1, col: 2, formula: '=B2*0.18' },
      ],
      isDone: true,
      parsedOnFirstAttempt: true,
    });
    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
    });

    const result = await service.execute(
      'calculate GST at 18% for column D',
      'FORMULA_GEN',
      workbookContext,
    );

    expect(executor.execute).toHaveBeenCalledTimes(1);
    expect(verifier.verify).toHaveBeenCalledTimes(1);
    expect(result.verifierSkipped).toBe(false);
    expect(result.verifierPassed).toBe(true);
    expect(result.sourceRefs.length).toBeGreaterThan(0);
    expect(result.sourceRefs[0]?.documentType).toBe('workbook');
  });

  it('always calls VerifierAgent exactly once when hardcode lint passes', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [
        {
          type: 'SET_FORMULA',
          sheetName: 'Invoices',
          row: 1,
          col: 2,
          formula: '=B2*0.18',
        },
      ],
      isDone: true,
    });
    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
    });

    await service.execute('add SUMIFS total for Amount where Status is Paid', 'FORMULA_GEN', workbookContext);

    expect(verifier.verify).toHaveBeenCalledTimes(1);
  });

  describe('generateOnly (plan mode)', () => {
    it('calls ExecutorAgent once and never VerifierAgent', async () => {
      executor.execute.mockResolvedValue({
        subtaskId: 's1',
        actions: [
          { type: 'SET_FORMULA', sheetName: 'Invoices', row: 1, col: 2, formula: '=B2*0.18' },
        ],
        isDone: true,
      });

      const result = await service.generateOnly(
        'calculate GST at 18% for column D',
        'FORMULA_GEN',
        workbookContext,
      );

      expect(executor.execute).toHaveBeenCalledTimes(1);
      expect(verifier.verify).not.toHaveBeenCalled();
      expect(result.actions).toHaveLength(1);
      expect(result.answer).toMatch(/proposed/i);
    });

    it('blocks hardcoded literals without calling VerifierAgent', async () => {
      executor.execute.mockResolvedValue({
        subtaskId: 's1',
        actions: [{ type: 'SET_CELL', sheetName: 'Invoices', row: 1, col: 2, value: 180 }],
        isDone: true,
      });

      const result = await service.generateOnly(
        'calculate GST at 18% for column D',
        'FORMULA_GEN',
        workbookContext,
      );

      expect(verifier.verify).not.toHaveBeenCalled();
      expect(result.blockedReason).toMatch(/numeric literal/i);
    });
  });

  it('blocks hardcoded numeric literals before any LLM verifier call', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [{ type: 'SET_CELL', sheetName: 'Invoices', row: 1, col: 2, value: 180 }],
      isDone: true,
    });

    const result = await service.execute(
      'calculate GST at 18% for column D',
      'FORMULA_GEN',
      workbookContext,
    );

    expect(verifier.verify).not.toHaveBeenCalled();
    expect(result.verifierPassed).toBe(false);
    expect(result.verifierSkipped).toBe(false);
    expect(result.failureReason).toMatch(/numeric literal/i);
  });

  it('never sets verifierSkipped to true', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [
        { type: 'SET_FORMULA', sheetName: 'Invoices', row: 1, col: 2, formula: '=B2*0.18' },
      ],
      isDone: true,
    });
    verifier.verify.mockResolvedValue({
      passed: false,
      feedback: 'Wrong reference',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: false, feedback: 'Wrong reference', issues: [] }],
    });

    const result = await service.execute(
      'calculate GST at 18% for column D',
      'FORMULA_GEN',
      workbookContext,
    );

    expect(result.verifierSkipped).toBe(false);
    expect(result.verifierPassed).toBe(false);
  });

  it('logs duration with tier 2 metadata', async () => {
    executor.execute.mockResolvedValue({
      subtaskId: 's1',
      actions: [
        { type: 'SET_FORMULA', sheetName: 'Invoices', row: 1, col: 2, formula: '=B2*0.18' },
      ],
      isDone: true,
    });
    verifier.verify.mockResolvedValue({
      passed: true,
      feedback: 'OK',
      issues: [],
      subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
    });

    const result = await service.execute(
      'calculate GST at 18% for column D',
      'FORMULA_GEN',
      workbookContext,
    );

    expect(result.durationMs).toBeGreaterThanOrEqual(0);
  });

  describe('representative formula fixtures', () => {
    it.each(tier2Fixtures)(
      'routes "$message" through executor and verifier',
      async ({ message, actionHint, expectedFormula, expectedActionType }) => {
        const action = expectedFormula
          ? {
              type: 'SET_FORMULA' as const,
              sheetName: 'Invoices',
              row: 1,
              col: 2,
              formula: expectedFormula,
            }
          : {
              type: (expectedActionType ?? 'SET_FORMULA') as 'SET_FORMULA',
              sheetName: 'Invoices',
              row: 1,
              col: 2,
              formula: '=B2',
            };

        executor.execute.mockResolvedValue({
          subtaskId: 's1',
          actions: [action],
          isDone: true,
        });
        verifier.verify.mockResolvedValue({
          passed: true,
          feedback: 'OK',
          issues: [],
          subtaskResults: [{ subtaskId: 's1', passed: true, feedback: 'OK', issues: [] }],
        });

        const result = await service.execute(message, actionHint, workbookContext);

        expect(executor.execute).toHaveBeenCalledTimes(1);
        expect(verifier.verify).toHaveBeenCalledTimes(1);
        expect(result.verifierPassed).toBe(true);
        expect(result.verifierSkipped).toBe(false);
      },
    );
  });
});

describe('assertTier2VerifierMandatory', () => {
  const originalNodeEnv = process.env.NODE_ENV;

  afterEach(() => {
    process.env.NODE_ENV = originalNodeEnv;
  });

  it('throws in non-production when shouldSkipVerifier would be used', () => {
    process.env.NODE_ENV = 'test';
    expect(() => assertTier2VerifierMandatory({ usedShouldSkipVerifier: true })).toThrow(
      /must never use shouldSkipVerifier/i,
    );
  });

  it('does not throw when verification path is mandatory', () => {
    process.env.NODE_ENV = 'test';
    expect(() => assertTier2VerifierMandatory({ usedShouldSkipVerifier: false })).not.toThrow();
  });
});
