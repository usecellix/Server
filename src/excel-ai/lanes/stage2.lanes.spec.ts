import {
  checkPanCell,
  computeDays,
  daysOutstandingLane,
  FAR_LAYOUT,
  farTemplateLane,
  fuzzyMatchLane,
  narrationLane,
  panLane,
  parseNarrationRow,
} from './stage2.lanes';
import { cellsOf, makeCtx, serial } from './lane-test-utils';

const T_FUZZY_A = [
  ['Vendor', 'Invoice'],
  ['ABC Pvt Ltd', 'INV/24-25/0233'],
  ['M/s Sharma Enterprises', 'INV-0101'],
  ['Kumar & Sons', '0450'],
  ['Delta Logistics Limited', 'DL/24-25/77'],
  ['Omega Steel Industries', 'OS-900'],
  ['Zenith Traders', 'Z-1'],
];
const T_FUZZY_B = [
  ['Vendor', 'Invoice'],
  ['ABC Private Limited', '233'],
  ['Sharma Enterprises', '101'],
  ['Kumar and Sons', '450'],
  ['Delta Logistic Ltd.', '77'],
  ['Omega Steels', 'OS900'],
  ['Quick Mart', 'Q-9'],
];
const T_DAYS = [
  ['Invoice', 'Party', 'Invoice Date', 'Amount', 'Credit Days'],
  ['INV-001', 'ABC Traders', serial(2026, 1, 15), 10000, 30],
  ['INV-002', 'XYZ Pvt Ltd', serial(2026, 2, 28), 25000, 45],
  ['INV-003', 'Mehta & Sons', serial(2026, 3, 31), 5000, 0],
  ['INV-004', 'Kumar Agency', serial(2024, 2, 29), 7000, 15],
  ['INV-005', 'Shah Bros', serial(2026, 4, 10), 3000, 30],
];

describe('P2.2 fuzzy match lane', () => {
  const msg = 'Match the vendor names in T_FuzzyA column A with T_FuzzyB column A';
  it('detects and ignores', () => {
    expect(fuzzyMatchLane.detect(msg)).toBe(true);
    expect(fuzzyMatchLane.detect('Reconcile purchase register with GSTR-2B')).toBe(false);
  });

  it('groups by band, Zenith unmatched, nothing linked outside the high band by default', async () => {
    const ctx = makeCtx(T_FUZZY_A, { message: msg, sheetName: 'T_FuzzyA' }, { T_FuzzyA: T_FUZZY_A, T_FuzzyB: T_FUZZY_B });
    const res = await fuzzyMatchLane.run(ctx);
    expect(res?.kind).toBe('actions');
    const answer = res!.answer;
    expect(answer).toMatch(/ABC Pvt Ltd ↔ ABC Private Limited/);
    expect(answer).toMatch(/Kumar & Sons ↔ Kumar and Sons/);
    expect(answer).toMatch(/Zenith Traders — no match/);
    const cells = cellsOf(res);
    expect(cells.C1).toBe('Matched in T_FuzzyB');
    expect(cells.C2).toBe('ABC Private Limited');
    expect(cells.C7).toBeUndefined(); // Zenith
    expect(Object.values(cells)).not.toContain('Review');
  });

  it('links the review band only when asked', async () => {
    const ctx = makeCtx(T_FUZZY_A, { message: `${msg} and also link the review band`, sheetName: 'T_FuzzyA' }, { T_FuzzyA: T_FUZZY_A, T_FuzzyB: T_FUZZY_B });
    const res = await fuzzyMatchLane.run(ctx);
    const answer = res!.answer;
    const reviewCount = Number(/\*\*Review[^*]*\*\* — (\d+)/.exec(answer)?.[1] ?? 0);
    const linkedReview = Object.values(cellsOf(res)).filter((v) => v === 'Review').length;
    expect(linkedReview).toBe(reviewCount);
  });
});

describe('P2.5 days outstanding lane', () => {
  const ctx = (message: string) => makeCtx(T_DAYS, { message });

  it('computes 75, 31, 0, 761, -10 and overdue 45, -14, 0, 746, -40', () => {
    const asOn = Date.UTC(2026, 2, 31);
    expect(computeDays(ctx(''), 2, asOn, -1).map((r) => r.days)).toEqual([75, 31, 0, 761, -10]);
    expect(computeDays(ctx(''), 2, asOn, 4).map((r) => r.overdue)).toEqual([45, -14, 0, 746, -40]);
  });

  it('asks for the as-on date instead of silently using today', () => {
    const res = daysOutstandingLane.run(ctx('Add a Days Outstanding column from Invoice Date'));
    expect(res).toMatchObject({ kind: 'answer' });
    expect((res as { answer: string }).answer).toMatch(/As on which date/);
  });

  it('writes live formulas with the explicit date and flags the future-dated row', () => {
    const res = daysOutstandingLane.run(ctx('Add a Days Outstanding column from Invoice Date as on 31-03-2026')) as never;
    const cells = cellsOf(res);
    expect(cells.F1).toBe('Days Outstanding');
    expect(cells.F2).toBe('=IF(ISNUMBER(C2),DATE(2026,3,31)-C2,"")');
    expect(cells.G1).toBe('Date Check');
    expect((res as { answer: string }).answer).toMatch(/75, 31, 0, 761, -10/);
    expect((res as { answer: string }).answer).toMatch(/future-dated row\(s\).*rows 6/);
  });

  it('adds overdue beyond credit days', () => {
    const res = daysOutstandingLane.run(ctx('Add days outstanding from Invoice Date as on 31-03-2026 and overdue beyond credit days')) as never;
    const cells = cellsOf(res);
    expect(cells.G1).toBe('Overdue Days');
    expect(cells.G2).toBe('=IF(ISNUMBER(C2),DATE(2026,3,31)-C2-N(E2),"")');
    expect((res as { answer: string }).answer).toMatch(/45, -14, 0, 746, -40/);
  });

  it('P3.11: writes into the named cell and shows a plain-English reading + 5 samples', () => {
    const res = daysOutstandingLane.run(ctx('Write a formula in F2 for days between Invoice Date and 31-03-2026')) as never;
    const cells = cellsOf(res);
    expect(cells.F2).toBe('=IF(ISNUMBER(C2),DATE(2026,3,31)-C2,"")');
    expect((res as { answer: string }).answer).toMatch(/Plain English/);
    expect((res as { answer: string }).answer).toMatch(/75, 31, 0, 761, -10/);
  });

  it('P3.14: verification covers blank, text, zero and negative rows', () => {
    const data = [...T_DAYS, ['INV-006', 'Blank', '', 1, 0], ['INV-007', 'Text', 'abc', 1, 0]];
    const res = daysOutstandingLane.run(makeCtx(data, { message: 'Write a formula for days outstanding in F2 as on 31-03-2026 and verify it' })) as never;
    const answer = (res as { answer: string }).answer;
    expect(answer).toMatch(/Verification/);
    expect(answer).toMatch(/Row 4: 0 \(same day\)/);
    expect(answer).toMatch(/Row 6: -10 \(future-dated/);
    expect(answer).toMatch(/Row 7: blank date/);
    expect(answer).toMatch(/Row 8: "abc" is not a date/);
  });
});

describe('P2.6 PAN lane', () => {
  it.each([
    ['AAAPA1234A', 'Valid', 'Individual'],
    ['AAACA1234A', 'Valid', 'Company'],
    ['AAAHA1234A', 'Valid', 'HUF'],
    ['AAAFA1234A', 'Valid', 'Firm / LLP'],
    ['AAATA1234A', 'Valid', 'Trust'],
    ['AAAPA12345', 'Invalid format', ''],
  ])('%s → %s %s', (raw, status, entity) => expect(checkPanCell(raw)).toMatchObject({ status, entity }));

  it('flags lowercase without correcting it', () => {
    expect(checkPanCell('aaapa1234a')).toMatchObject({ pan: 'aaapa1234a', status: expect.stringMatching(/lowercase/) });
  });

  it('extracts a PAN from a sentence', () => {
    expect(checkPanCell('Vendor ABC PAN AABCA1234F paid on 5th')).toMatchObject({ pan: 'AABCA1234F', status: 'Valid (extracted)', entity: 'Company' });
  });

  it('adds three columns without touching column A', () => {
    const vals = [['Text'], ['AAAPA1234A'], ['aaapa1234a']];
    const res = panLane.run(makeCtx(vals, { message: 'Validate the PANs in column A and show the entity type' })) as never;
    const cells = cellsOf(res);
    expect(cells.B1).toBe('PAN');
    expect(cells.D2).toBe('Individual');
    expect(Object.keys(cells).some((a) => a.startsWith('A'))).toBe(false);
  });
});

describe('P2.7 FAR template lane', () => {
  it('creates a uniquely named sheet with the frozen layout', () => {
    expect(farTemplateLane.detect('Create a fixed asset register')).toBe(true);
    const res = farTemplateLane.run(makeCtx([[]])) as unknown as { kind: string; actions: Array<Record<string, unknown>> };
    expect(res.actions[0]).toMatchObject({ type: 'CREATE_SHEET', uniqueName: true });
    expect(res.actions[1]).toMatchObject({ type: 'WRITE_TABLE', headers: [...FAR_LAYOUT] });
    expect(FAR_LAYOUT).toEqual(['Asset', 'Block/Class', 'Date of Purchase', 'Put-to-use Date', 'Cost', 'Additions', 'Deletions', 'Location', 'Tag No.']);
    expect(res.actions.some((a) => (a.format as { numberFormat?: string } | undefined)?.numberFormat === 'dd-mm-yyyy')).toBe(true);
  });
});

describe('P2.8 narration lane', () => {
  it.each([
    ['UPI/412345678901/Payment to ABC TRADERS/abc@okhdfc/HDFC', 'UPI', '412345678901', 'ABC TRADERS'],
    ['NEFT-SBIN0001234-XYZ PVT LTD-INV 233-SBIN52026010112345', 'NEFT', 'SBIN52026010112345', 'XYZ PVT LTD'],
    ['RTGS/HDFCR52026011500123/MEHTA & SONS', 'RTGS', 'HDFCR52026011500123', 'MEHTA & SONS'],
    ['IMPS/P2A/412398765432/KUMAR AGENCY', 'IMPS', '412398765432', 'KUMAR AGENCY'],
    ['CHQ DEP 000123 SHAH BROS', 'CHQ', '000123', 'SHAH BROS'],
  ])('%s', (raw, mode, reference, counterparty) => {
    expect(parseNarrationRow(raw)).toMatchObject({ mode, reference, counterparty, parsed: true });
  });

  it('extracts the invoice number from the NEFT line', () => {
    expect(parseNarrationRow('NEFT-SBIN0001234-XYZ PVT LTD-INV 233-SBIN52026010112345').invoice).toBe('233');
  });

  it('lists gibberish as unparsed instead of guessing', () => {
    const vals = [['Narration'], ['UPI/412345678901/Payment to ABC TRADERS/abc@okhdfc/HDFC'], ['gibberish text 1234']];
    const res = narrationLane.run(makeCtx(vals, { message: 'Parse the bank narrations in column A' })) as never;
    expect((res as { answer: string }).answer).toMatch(/1 not recognised.*\n- row 3 — gibberish text 1234/);
    expect(cellsOf(res).B3).toBe('Unparsed');
  });
});

describe('P4.8 bank packs from the rule table', () => {
  it('a bank-specific prefix the generic parser misses is recognised from the pack', () => {
    // "BY CLEARING" (SBI cheque clearing) has no CHQ/NEFT keyword.
    expect(parseNarrationRow('BY CLEARING 456789 RAJ & CO').mode).toBe('CHQ');
  });
});
