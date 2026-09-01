import {
  annotateAnswerConsistency,
  checkAnswerConsistency,
  extractLabelledAmounts,
} from '../src/excel-ai/utils/answer-consistency.util';

/** Verbatim from the task #90 incident (Aug 27, 2026). */
const REAL_ANSWER = `Purchase Register overview
- 50 records, 01-01-2026 – 19-02-2026
- Columns: Date, Invoice No, Supplier, GSTIN, Item, Qty, Unit Price, Tax %, Tax Amount, Total Amount, Payment Status, Remarks
- Total quantity: 275 units · Pre-tax: ₹14,275 · GST: ₹14,148 · Total: ₹92,748

Payment Status
- Paid: 33 records, ₹60,407
- Pending: 17 records, ₹32,341`;

/** The same answer with the pre-tax figure corrected to the true value. */
const CORRECT_ANSWER = REAL_ANSWER.replace('Pre-tax: ₹14,275', 'Pre-tax: ₹78,600');

describe('extractLabelledAmounts', () => {
  it('pulls labelled figures out of the middot-separated layout the model produces', () => {
    const amounts = extractLabelledAmounts(REAL_ANSWER);
    const byLabel = Object.fromEntries(amounts.map((a) => [a.label, a.value]));
    expect(byLabel['Pre-tax']).toBe(14275);
    expect(byLabel['GST']).toBe(14148);
    expect(byLabel['Total']).toBe(92748);
  });

  it('strips currency symbols and thousands separators', () => {
    const amounts = extractLabelledAmounts('Net: ₹1,23,456.50');
    expect(amounts[0]?.value).toBe(123456.5);
  });
});

describe('checkAnswerConsistency', () => {
  it('flags the real incident: pre-tax + GST does not equal the stated total', () => {
    const issue = checkAnswerConsistency(REAL_ANSWER);
    expect(issue).not.toBeNull();
    expect(issue!.kind).toBe('subtotal_mismatch');
    expect(issue!.expected).toBe(14275 + 14148);
    expect(issue!.stated).toBe(92748);
  });

  it('passes the same answer once the wrong figure is corrected', () => {
    expect(checkAnswerConsistency(CORRECT_ANSWER)).toBeNull();
  });

  it('tolerates rounding differences', () => {
    expect(
      checkAnswerConsistency('Pre-tax: 78600.40 · GST: 14148.10 · Total: 92748.00'),
    ).toBeNull();
  });

  it('does not fire when the answer states too little to check', () => {
    expect(checkAnswerConsistency('Total: ₹92,748')).toBeNull();
    expect(checkAnswerConsistency('There are 50 rows in this sheet.')).toBeNull();
    expect(checkAnswerConsistency('')).toBeNull();
  });

  it('does not confuse "Total tax" with the grand total', () => {
    // "Total tax collected" must be read as the TAX figure, not the total.
    const answer = 'Taxable value: 78,600 · Total tax collected: 14,148 · Grand total: 92,748';
    expect(checkAnswerConsistency(answer)).toBeNull();
  });

  it('recognises subtotal/net wording as the pre-tax figure', () => {
    expect(checkAnswerConsistency('Subtotal: 100 · Tax: 18 · Total: 118')).toBeNull();
    expect(checkAnswerConsistency('Subtotal: 100 · Tax: 18 · Total: 500')).not.toBeNull();
  });
});

describe('annotateAnswerConsistency', () => {
  it('appends a verification warning to the real incident answer', () => {
    const { answer, issue } = annotateAnswerConsistency(REAL_ANSWER);
    expect(issue).not.toBeNull();
    expect(answer).toMatch(/Please verify these figures/i);
    expect(answer).toContain(REAL_ANSWER); // original content preserved
  });

  it('does not guess which figure is wrong', () => {
    // The true pre-tax (78,600) is NOT asserted — we only report the contradiction,
    // so the user is never handed a second unverified number.
    const { answer } = annotateAnswerConsistency(REAL_ANSWER);
    expect(answer).not.toContain('78,600');
  });

  it('leaves a consistent answer untouched', () => {
    const { answer, issue } = annotateAnswerConsistency(CORRECT_ANSWER);
    expect(issue).toBeNull();
    expect(answer).toBe(CORRECT_ANSWER);
  });
});
