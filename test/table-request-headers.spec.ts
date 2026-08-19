import {
  buildTableActionsFromMessage,
  parseTableCreateRequest,
  tryDeterministicTableCreate,
} from '../src/excel-ai/utils/table-request.util';

describe('table-request.util — add headers + sample rows', () => {
  const prompt =
    'add headers Job Title, Company, Student Name, Student Email, Status and 3 sample rows';

  it('parses the user test-script prompt into WRITE_TABLE', () => {
    const plan = parseTableCreateRequest(prompt);
    expect(plan).not.toBeNull();
    expect(plan!.headers).toEqual([
      'Job Title',
      'Company',
      'Student Name',
      'Student Email',
      'Status',
    ]);
    expect(plan!.rowCount).toBe(3);
    expect(plan!.rows).toHaveLength(3);
    expect(plan!.rows[0]).toHaveLength(5);
  });

  it('builds a single WRITE_TABLE action (deterministic path)', () => {
    const actions = buildTableActionsFromMessage(prompt);
    expect(actions).toHaveLength(1);
    expect(actions![0].type).toBe('WRITE_TABLE');
    expect(actions![0].headers).toContain('Student Email');
    expect(actions![0].rows).toHaveLength(3);
  });

  it('generates email-like sample values for Student Email', () => {
    const plan = parseTableCreateRequest(prompt)!;
    const emailCol = plan.headers.indexOf('Student Email');
    expect(String(plan.rows[0][emailCol])).toMatch(/@example\.com$/);
  });
});

describe('table-request.util — compound purchase register prompt (bug regression)', () => {
  const compoundPrompt =
    'Create a purchase register from the data in this workbook. Add columns for purchase date, ' +
    'supplier, invoice number, item, category, quantity, unit price, tax %, tax amount, total amount, ' +
    'payment status, department, requested by, and approved by. Add formulas for tax and total amount. ' +
    'Add filters, freeze the header row, and create a summary showing total purchases, paid amount, ' +
    'pending amount, and purchases by department.';

  it('does not extract "header row" as a header list and returns no deterministic plan', () => {
    const plan = parseTableCreateRequest(compoundPrompt);
    expect(plan).toBeNull();
  });
});

describe('tryDeterministicTableCreate — single gate for both call sites (Phase D)', () => {
  it('rejects a compound prompt that would otherwise build a garbage table', () => {
    // Same prompt as the bug-regression case above: parseTableCreateRequest
    // alone already returns null for it, but this exercises the gate that
    // conversation.service.ts actually calls from both the pre-LLM path and
    // the LLM-parse-failure fallback path.
    const compoundPrompt =
      'Create a purchase register from the data in this workbook. Add columns for purchase date, ' +
      'supplier, invoice number, item, category, quantity, unit price, tax %, tax amount, total amount, ' +
      'payment status, department, requested by, and approved by. Add formulas for tax and total amount. ' +
      'Add filters, freeze the header row, and create a summary showing total purchases, paid amount, ' +
      'pending amount, and purchases by department.';
    expect(tryDeterministicTableCreate(compoundPrompt)).toBeNull();
  });

  it('rejects a compound prompt even when phrased without the "header row" false match', () => {
    // Regression target: before this gate was shared, the LLM-parse-failure
    // fallback had no hasCompoundSignals check at all, so a compound prompt
    // that dodges extractHeaders' specific false-match (no literal "header
    // row" text) could still slip through as a deterministic table create,
    // silently dropping every clause after the first.
    const compoundPrompt =
      'Add columns Name, Age, City, and then add a chart of the data.';
    expect(tryDeterministicTableCreate(compoundPrompt)).toBeNull();
  });

  it('still builds a table for a genuinely simple, non-compound request', () => {
    const result = tryDeterministicTableCreate(
      'add headers Job Title, Company, Student Name and 3 sample rows',
    );
    expect(result).not.toBeNull();
    expect(result!.plan.headers).toEqual(['Job Title', 'Company', 'Student Name']);
    expect(result!.actions).toHaveLength(1);
    expect(result!.actions[0].type).toBe('WRITE_TABLE');
  });
});
