import {
  buildTableActionsFromMessage,
  parseTableCreateRequest,
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
