import { isReadOnlyQuestion } from '../src/excel-ai/services/llm-router.service';

/**
 * TASKS.md #214 — `cellix-basic-usecases.html` Q&A.4 ("Data Analysis
 * Questions") is explicitly read-only: "compute and return the answer without
 * writing to the sheet". The live audit caught "Are there duplicate values in
 * column A?" being routed to write, where Tier 2 inserted a `Duplicate?`
 * column into the user's data to answer a yes/no question.
 */
describe('read-only question detection (#214)', () => {
  it.each([
    'Are there duplicate values in column A?',
    'How many blank cells are in the GSTIN column?',
    'Which supplier has the highest total taxable amount?',
    'What is the total purchase amount from Kerala suppliers?',
    'How many rows of data do I have?',
    'What is in cell E3?',
    'Do I have any rows with a missing GSTIN?',
    'What filter is active right now?',
  ])('treats %j as read-only', (message) => {
    expect(isReadOnlyQuestion(message)).toBe(true);
  });

  it.each([
    'Can you highlight the duplicates in column A?',
    'Could you add a Status column?',
    'How about you sort this by date?',
    'Highlight duplicate invoice numbers in column A',
    'Which rows are duplicates? Delete them',
    'What if you format column E as currency?',
  ])('leaves %j as a write request', (message) => {
    expect(isReadOnlyQuestion(message)).toBe(false);
  });
});
