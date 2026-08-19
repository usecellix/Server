import { hasCompoundSignals } from '../src/excel-ai/utils/complexity-classifier.util';

describe('hasCompoundSignals — gates the deterministic table-shortcut', () => {
  it('is true for the purchase-register bug prompt (compound, multi-feature)', () => {
    const prompt =
      'Create a purchase register from the data in this workbook. Add columns for purchase date, ' +
      'supplier, invoice number, item, category, quantity, unit price, tax %, tax amount, total amount, ' +
      'payment status, department, requested by, and approved by. Add formulas for tax and total amount. ' +
      'Add filters, freeze the header row, and create a summary showing total purchases, paid amount, ' +
      'pending amount, and purchases by department.';
    expect(hasCompoundSignals(prompt)).toBe(true);
  });

  it('is false for a genuinely simple table-create request', () => {
    expect(
      hasCompoundSignals('add headers Job Title, Company, Student Name and 3 sample rows'),
    ).toBe(false);
  });

  it('is true for an Oxford-comma header list (accepted trade-off: falls to the LLM path, still correct)', () => {
    expect(hasCompoundSignals('create a table with columns Name, Age, and City')).toBe(true);
  });
});
