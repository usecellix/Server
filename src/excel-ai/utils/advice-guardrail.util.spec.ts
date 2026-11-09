import {
  ADVICE_DISCLAIMER,
  applyAdviceGuardrail,
  containsVerdictLanguage,
  isAdviceShapedQuestion,
} from './advice-guardrail.util';

describe('advice guardrail (P0.2)', () => {
  describe('isAdviceShapedQuestion', () => {
    it.each([
      'Can I claim ITC on this invoice?',
      'Should I file GSTR-3B now?',
      'Is this deductible under 43B(h)?',
      'am I eligible for set off',
      'Is it allowed to claim this after 30 Nov?',
      'Will I get a penalty for late TDS deposit',
    ])('flags %s', (q) => {
      expect(isAdviceShapedQuestion(q)).toBe(true);
    });

    it.each([
      'Total of column D',
      'Trim spaces in column A',
      'Reconcile purchase register with GSTR-2B',
      'Fill A1:A5 with 1',
      '',
    ])('does not flag %s', (q) => {
      expect(isAdviceShapedQuestion(q)).toBe(false);
    });
  });

  describe('applyAdviceGuardrail', () => {
    it('appends the fixed disclaimer for an advice-shaped question', () => {
      const r = applyAdviceGuardrail('Invoice 233 has ITC of ₹18,000; 2 of 5 invoices are beyond the Sec 16(4) date.', 'Can I claim ITC on this invoice?');
      expect(r.guarded).toBe(true);
      expect(r.answer).toContain(ADVICE_DISCLAIMER);
      expect(r.answer).toContain('₹18,000'); // facts preserved
    });

    it('removes verdict wording even if the question was not advice-shaped', () => {
      const r = applyAdviceGuardrail('You can claim this ITC. Also you should file now.', 'summarise the sheet');
      expect(r.guarded).toBe(true);
      expect(r.answer).not.toMatch(/you can claim/i);
      expect(r.answer).not.toMatch(/you should file/i);
      expect(containsVerdictLanguage(r.answer)).toBe(false);
      expect(r.answer).toContain(ADVICE_DISCLAIMER);
    });

    it('softens instruction-style filing guidance seen in live testing', () => {
      const live =
        "I can't give a definitive yes. As a general rule, file by the applicable due date once data is reconciled, and avoid filing if you suspect errors.";
      const r = applyAdviceGuardrail(live, 'Should I file GSTR-3B now?');
      expect(r.answer).not.toMatch(/file by the/i);
      expect(r.answer).not.toMatch(/avoid filing/i);
      expect(r.answer).not.toMatch(/as a general rule/i);
      expect(r.answer).toContain(ADVICE_DISCLAIMER);
    });

    it('leaves plain computations untouched', () => {
      const r = applyAdviceGuardrail('The total for Amount is ₹12,34,567.', 'total of column D');
      expect(r).toEqual({ answer: 'The total for Amount is ₹12,34,567.', guarded: false });
    });

    it('is idempotent (disclaimer never doubled)', () => {
      const once = applyAdviceGuardrail('Fact.', 'Should I file GSTR-3B now?').answer;
      const twice = applyAdviceGuardrail(once, 'Should I file GSTR-3B now?').answer;
      expect(twice.split(ADVICE_DISCLAIMER).length - 1).toBe(1);
    });

    it('handles empty / non-string answers safely', () => {
      expect(applyAdviceGuardrail('', 'Can I claim this?')).toEqual({ answer: '', guarded: false });
    });
  });
});
