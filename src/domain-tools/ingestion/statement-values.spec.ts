import {
  inferNumericDateOrder,
  isoDateToExcelSerial,
  looksLikeDate,
  parseStatementAmount,
  parseStatementDate,
  toIsoDate,
  tokenKind,
} from './statement-values';

const iso = (text: string, order?: 'dmy' | 'mdy') => {
  const parts = parseStatementDate(text, order);
  return parts ? toIsoDate(parts) : null;
};

describe('parseStatementDate', () => {
  it.each([
    ['11-APR-2026', '2026-04-11'],
    ['11 Apr 2026', '2026-04-11'],
    ['11Apr26', '2026-04-11'],
    ['1-Sept-2026', '2026-09-01'],
    ['Apr 11, 2026', '2026-04-11'],
    ['2026-04-11', '2026-04-11'],
    ['11/04/2026', '2026-04-11'],
    ['11.04.26', '2026-04-11'],
    ['11/04/2026 14:32:10', '2026-04-11'],
  ])('reads %s', (text, expected) => {
    expect(iso(text)).toBe(expected);
  });

  it('reads an all-numeric date in the order it is told', () => {
    expect(iso('04/05/2026', 'dmy')).toBe('2026-05-04');
    expect(iso('04/05/2026', 'mdy')).toBe('2026-04-05');
  });

  it.each(['', 'Opening Balance', '31/02/2026', '12ABC34', 'S93580561', '1627.00', '00-APR-2026'])(
    'rejects %p',
    (text) => {
      expect(iso(text)).toBeNull();
    },
  );

  it('recognises a date that is only valid month-first', () => {
    expect(iso('04/25/2026')).toBeNull();
    expect(looksLikeDate('04/25/2026')).toBe(true);
  });
});

describe('inferNumericDateOrder', () => {
  it('is day-first when a first part exceeds 12', () => {
    expect(inferNumericDateOrder(['01/04/2026', '13/04/2026'])).toBe('dmy');
  });

  it('is month-first when a second part exceeds 12', () => {
    expect(inferNumericDateOrder(['04/01/2026', '04/13/2026'])).toBe('mdy');
  });

  it('picks the reading that keeps an ambiguous column in order', () => {
    // Both readings are in order here, so this one falls back to day-first.
    expect(inferNumericDateOrder(['01/02/2026', '01/03/2026', '01/04/2026', '01/05/2026'])).toBe('dmy');
    // Day-first: 1 Feb, 2 Feb, 1 Mar, 2 Mar, 1 Apr (in order). Month-first: Jan 2, Feb 2, Jan 3, Feb 3, Jan 4 (zig-zag).
    expect(
      inferNumericDateOrder(['01/02/2026', '02/02/2026', '01/03/2026', '02/03/2026', '01/04/2026']),
    ).toBe('dmy');
    // Month-first: Feb 1, Feb 2, Mar 1, Mar 2, Apr 1 (in order). Day-first zig-zags.
    expect(
      inferNumericDateOrder(['02/01/2026', '02/02/2026', '03/01/2026', '03/02/2026', '04/01/2026']),
    ).toBe('mdy');
  });

  it('defaults to day-first when nothing settles it', () => {
    expect(inferNumericDateOrder(['01/02/2026'])).toBe('dmy');
    expect(inferNumericDateOrder(['11-APR-2026', '12-APR-2026'])).toBe('dmy');
  });
});

describe('isoDateToExcelSerial', () => {
  it('matches Excel for known dates', () => {
    expect(isoDateToExcelSerial('2000-01-01')).toBe(36526);
    expect(isoDateToExcelSerial('2026-04-11')).toBe(46123);
    expect(isoDateToExcelSerial('2024-02-29')).toBe(45351);
  });

  it('returns null for anything that is not an ISO date', () => {
    expect(isoDateToExcelSerial('')).toBeNull();
    expect(isoDateToExcelSerial('11-APR-2026')).toBeNull();
  });
});

describe('parseStatementAmount', () => {
  it.each([
    ['1627.00', 162700],
    ['1,25,000.00', 12500000],
    ['125,000.50', 12500050],
    ['₹ 1,000', 100000],
    ['Rs. 250.5', 25050],
    ['INR 99.99', 9999],
    ['0.00', 0],
    ['.50', 50],
    ['12.345', 1235],
    ['105000.00', 10500000],
  ])('reads %s', (text, paise) => {
    expect(parseStatementAmount(text)).toEqual({ paise, negative: false });
  });

  it.each(['(500.00)', '-500.00', '500.00-', '− 500.00'])('reads %s as negative', (text) => {
    expect(parseStatementAmount(text)).toEqual({ paise: 50000, negative: true });
  });

  it.each([
    ['1,234.00 Cr', 'cr'],
    ['1234.00Cr', 'cr'],
    ['1,234.00 CR.', 'cr'],
    ['1,234.00 (Dr)', 'dr'],
    ['Dr 1,234.00', 'dr'],
    ['1,234.00 Debit', 'dr'],
  ])('reads the Dr/Cr marker in %s', (text, marker) => {
    expect(parseStatementAmount(text)).toEqual({ paise: 123400, negative: false, marker });
  });

  it.each(['', '   ', '-', '--', 'NIL', 'n/a', 'Cr', 'TFR', 'S93580561', '12-04-2026', '1.2.3'])(
    'returns null for %p',
    (text) => {
      expect(parseStatementAmount(text)).toBeNull();
    },
  );
});

describe('tokenKind', () => {
  it.each([
    ['11-APR-2026', 'date'],
    ['11/04/2026', 'date'],
    ['1627.00', 'money'],
    ['1,25,000', 'money'],
    ['105000.00', 'money'],
    ['004521', 'integer'],
    ['28', 'integer'],
    ['S93580561', 'text'],
    ['Cr', 'text'],
    ['UPIOUT/631653934075', 'text'],
    ['192.168.1.1', 'text'],
    ['', 'text'],
  ])('classifies %p as %s', (text, kind) => {
    expect(tokenKind(text)).toBe(kind);
  });
});
