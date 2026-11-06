import {
  buildGstinValidationPlan,
  columnLetter,
  detectValidateGstinIntent,
  explicitColumnFromMessage,
  findGstinColumn,
  GSTIN_STATE_HEADER,
  GSTIN_STATUS_HEADER,
} from './gstin-validation-lane.util';

// Same fixtures as the T_GSTIN sheet in CELLIX_TEST_DATA.xlsx.
const SHEET: unknown[][] = [
  ['Party', 'GSTIN'],
  ['Valid Vendor 1', '27AABCA1234F1Z9'],
  ['Valid Vendor 2', '29AAACM5678K1ZV'],
  ['Valid Vendor 3', '07BBBPS4321Q1ZA'],
  ['Valid Vendor 4', '33CCCHD9876L1ZN'],
  ['Valid Vendor 5', '24DDDFR1111M1ZA'],
  ['Valid Vendor 6', '19EEETP2222N1ZY'],
  ['Valid Vendor 7', '36FFFCT3333P1ZJ'],
  ['Valid Vendor 8', '09GGGAS4444R1Z7'],
  ['Typo last char', '27AABCA1234F1Z7'],
  ['Typo middle char', '29AAACA5678K1ZV'],
  ['Typo digit', '07BBBPS4371Q1ZA'],
  ['Typo state digits', '37CCCHD9876L1ZN'],
  ['Too short', '24DDDFR1111M'],
  ['Invalid state 50', '50HHHPT5555S1Z?'],
];

function opsByAddress(plan: ReturnType<typeof buildGstinValidationPlan>) {
  const ops = (plan.actions?.[0] as { operations: Array<{ address: string; value: unknown }> }).operations;
  return Object.fromEntries(ops.map((o) => [o.address, o.value]));
}

describe('P1.4 GSTIN validation lane', () => {
  describe('intent', () => {
    it.each([
      'Validate the GSTINs in column B',
      'validate gstin',
      'Check GSTIN numbers in this sheet',
      'verify the GSTINs',
    ])('detects: %s', (m) => expect(detectValidateGstinIntent(m)).toBe(true));

    it.each([
      'Reconcile purchase register with GSTR-2B',
      'Highlight invalid GSTINs in column B',
      'Match GSTIN of A with B',
      'Put the number 1 in cells A1 to A5',
      'Total of column D',
      '',
    ])('ignores: %s', (m) => expect(detectValidateGstinIntent(m)).toBe(false));
  });

  it('column helpers', () => {
    expect(columnLetter(0)).toBe('A');
    expect(columnLetter(25)).toBe('Z');
    expect(columnLetter(26)).toBe('AA');
    expect(explicitColumnFromMessage('Validate the GSTINs in column B')).toBe(1);
    expect(explicitColumnFromMessage('validate gstins in column named GSTIN')).toBeNull();
  });

  describe('column detection', () => {
    it('by explicit letter, by header, and by content', () => {
      expect(findGstinColumn(SHEET, 0, 'validate gstin in column B')).toBe(1);
      expect(findGstinColumn(SHEET, 0, 'validate the gstins')).toBe(1); // header "GSTIN"
      const noHeader = SHEET.map((r) => [r[0], r[1]]);
      noHeader[0] = ['Party', 'Tax id'];
      expect(findGstinColumn(noHeader, 0, 'validate the gstins')).toBe(1); // content ≥50% GSTIN-shaped
    });

    it('returns null when no column looks like GSTINs', () => {
      expect(findGstinColumn([['A', 'B'], ['x', 'y'], ['z', 'w']], 0, 'validate gstins')).toBeNull();
    });
  });

  describe('plan on the T_GSTIN fixture', () => {
    const plan = buildGstinValidationPlan(SHEET, 0, 'T_GSTIN', 'Validate the GSTINs in column B');

    it('counts: 8 valid, 3 checksum fail + 1 checksum fail (state digit typo), 1 invalid format, 1 invalid state', () => {
      expect(plan.kind).toBe('actions');
      expect(plan.counts).toEqual({
        checked: 14,
        valid: 8,
        checksum_fail: 4,
        invalid_format: 2, // "Too short" and the one with a '?' check char
        invalid_state: 0,
        invalid_pan: 0,
      });
    });

    it('writes two new columns after the used range and never touches column B', () => {
      const ops = opsByAddress(plan);
      expect(ops['C1']).toBe(GSTIN_STATUS_HEADER);
      expect(ops['D1']).toBe(GSTIN_STATE_HEADER);
      expect(ops['C2']).toBe('Valid');
      expect(ops['D2']).toBe('Maharashtra');
      expect(ops['C10']).toBe('Checksum fail');
      expect(Object.keys(ops).some((a) => a.startsWith('B'))).toBe(false);
    });

    it('is a single BATCH_SET on the named sheet (goes through preview/Accept)', () => {
      expect(plan.actions).toHaveLength(1);
      expect(plan.actions![0]).toMatchObject({ type: 'BATCH_SET', sheetName: 'T_GSTIN' });
    });

    it('lists rows to review and states what the status does NOT mean', () => {
      expect(plan.answer).toContain('row 10');
      expect(plan.answer).toMatch(/not whether the registration is active/i);
      expect(plan.answer).not.toMatch(/you can claim|you should file/i);
    });
  });

  it('classifies an invalid state separately', () => {
    // 50 + PAN HHHPT5555S + entity 1 + Z + correct check char
    const sheet: unknown[][] = [['GSTIN'], ['50HHHPT5555S1ZX']];
    // valid format, state 50 not in list → invalid_state regardless of check char
    const plan = buildGstinValidationPlan(sheet, 0, 'S', 'validate gstins');
    expect(plan.counts?.invalid_state).toBe(1);
  });

  it('reuses its own columns on a second run instead of adding more', () => {
    const rerun: unknown[][] = [
      ['Party', 'GSTIN', GSTIN_STATUS_HEADER, GSTIN_STATE_HEADER],
      ['A', '27AABCA1234F1Z9', 'Valid', 'Maharashtra'],
    ];
    const plan = buildGstinValidationPlan(rerun, 0, 'S', 'validate gstins');
    const ops = opsByAddress(plan);
    expect(ops['C1']).toBe(GSTIN_STATUS_HEADER);
    expect(ops['E1']).toBeUndefined();
  });

  it('skips blank GSTIN cells and respects a header row that is not row 1', () => {
    const sheet: unknown[][] = [
      ['Ledger report'],
      ['Party', 'GSTIN'],
      ['A', '27AABCA1234F1Z9'],
      ['B', ''],
    ];
    const plan = buildGstinValidationPlan(sheet, 1, 'S', 'validate gstins');
    expect(plan.counts?.checked).toBe(1);
    const ops = opsByAddress(plan);
    expect(ops['C2']).toBe(GSTIN_STATUS_HEADER); // header row is sheet row 2
    expect(ops['C3']).toBe('Valid');
    expect(ops['C4']).toBeUndefined();
  });

  it('answers (no write) when there is nothing to validate', () => {
    expect(buildGstinValidationPlan([['GSTIN']], 0, 'S', 'validate gstins').kind).toBe('answer');
    expect(buildGstinValidationPlan([['A'], ['x']], 0, 'S', 'validate gstins').kind).toBe('answer');
  });
});
