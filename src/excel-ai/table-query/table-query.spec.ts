import { buildTable, isDateNumberFormat, resolveColumn, serialToIso, TableModel, toNumber } from './table-model';
import { executeTableQuery, sumExact } from './table-query.executor';
import { formatDate, formatNumber, formatTableAnswer } from './table-query.format';
import { parseTableQueryPlan } from './table-query.plan';
import { buildTableQueryUserMessage, commonStarts, looksLikeTableQuestion, TableQuestionAnswerer } from './table-question';
import { TableQueryPlanError } from './table-query.types';

/** A small statement as Excel hands it over: dates as serials, blanks as '', a header row. */
const HEADERS = ['Date', 'Value Date', 'Description', 'Ref No', 'Debit', 'Credit', 'Balance', 'Source', 'Flag'];
const D = (iso: string) => Math.round(Date.parse(`${iso}T00:00:00Z`) / 86_400_000) + 25_569;
const ROWS: unknown[][] = [
  [D('2026-04-01'), D('2026-04-01'), 'NEFT CR-SAMPLE EMPLOYER LTD-SALARY', 'N1', '', 50000, 75000, 'row 2', ''],
  [D('2026-04-03'), D('2026-04-03'), 'ATM WDL 593111 SAMPLE TOWN', '785686', 6173.55, '', 68826.45, 'row 3', ''],
  [D('2026-04-06'), D('2026-04-06'), 'UPI-GROCERY MART-payee6@testupi', '4840', 1925.41, '', 66901.04, 'row 4', ''],
  [D('2026-04-28'), D('2026-04-28'), 'CHQ PAID-MICR CTS-BOOK HOUSE', '006076', 15000, '', 51901.04, 'row 5', ''],
  [D('2026-05-06'), D('2026-05-07'), 'ATM WDL 605256 SAMPLE TOWN', '105457', 15000, '', 36901.04, 'row 6', ''],
  [D('2026-05-08'), D('2026-05-08'), 'CHQ PAID-MICR CTS-METRO CAFE', '005736', 15000, '', 21901.04, 'row 7', ''],
  [D('2026-05-09'), D('2026-05-09'), 'INT.PD:01-01-2026 TO 31-03-2026', 'I9', '', 220.78, 22121.82, 'row 8', ''],
  [D('2026-05-18'), D('2026-05-18'), 'CHQ PAID-MICR CTS-FUEL STATION', '001859', 20000, '', 2121.82, 'row 9', ''],
];
const HINTS = [
  { index: 0, detectedType: 'date', numberFormat: 'dd-mmm-yyyy' },
  { index: 1, detectedType: 'date', numberFormat: 'dd-mmm-yyyy' },
  { index: 4, detectedType: 'number', numberFormat: '#,##0.00' },
  { index: 5, detectedType: 'number', numberFormat: '#,##0.00' },
  { index: 6, detectedType: 'number', numberFormat: '#,##0.00' },
  // The add-in reports an all-blank column as a number. It must still read as text.
  { index: 8, detectedType: 'number' },
];
const table = (): TableModel => buildTable({ sheetName: 'Bank Statement', sheetData: [HEADERS, ...ROWS], columnHints: HINTS });
const run = (plan: unknown, t = table()) => {
  const parsed = parseTableQueryPlan(t, JSON.stringify(plan));
  if (parsed.kind !== 'queries') throw new Error('expected queries');
  return parsed.queries.map((q) => executeTableQuery(t, q));
};
const answer = (plan: unknown) => formatTableAnswer(table(), run(plan));

describe('buildTable', () => {
  it('types the columns of an imported statement', () => {
    expect(table().columns.map((c) => [c.name, c.type, c.decimals])).toEqual([
      ['Date', 'date', 0], ['Value Date', 'date', 0], ['Description', 'text', 0], ['Ref No', 'text', 0],
      ['Debit', 'number', 2], ['Credit', 'number', 2], ['Balance', 'number', 2], ['Source', 'text', 0], ['Flag', 'text', 0],
    ]);
  });

  it('keeps the sheet row number of every row, and skips blank rows', () => {
    const t = buildTable({ sheetName: 'S', sheetData: [['Title'], ['A', 'B'], [1, 2], ['', ''], [3, 4]], headerRowIndex: 1 });
    expect(t.rows.map((r) => r.rowNumber)).toEqual([3, 5]);
    expect(t.columns.map((c) => c.name)).toEqual(['A', 'B']);
  });

  it('does not take a plain number column for dates', () => {
    const t = buildTable({ sheetName: 'S', sheetData: [['Amount'], [45000], [46113], [52000]] });
    expect(t.columns[0].type).toBe('number');
  });

  it('reads dates typed as text', () => {
    const t = buildTable({ sheetName: 'S', sheetData: [['When'], ['11/04/2026'], ['12-Apr-2026']] });
    expect(t.columns[0].type).toBe('date');
  });

  it('names blank and repeated headers so every column can be addressed', () => {
    const t = buildTable({ sheetName: 'S', sheetData: [['Amount', '', 'Amount'], [1, 2, 3]] });
    expect(t.columns.map((c) => c.name)).toEqual(['Amount', 'Column 2', 'Amount (2)']);
  });
});

describe('cell reading', () => {
  it.each([[1250.5, 1250.5], ['1,25,000.50', 125000.5], ['₹ 500', 500], ['(250)', -250], ['-12.5', -12.5], ['12%', 12]])(
    'reads %p as a number', (raw, n) => expect(toNumber(raw)).toBe(n),
  );
  it.each(['', 'ATM', '12-04-2026', 'S123', null])('does not read %p as a number', (raw) => expect(toNumber(raw)).toBeNull());

  it('converts Excel serials to dates', () => {
    expect(serialToIso(46113)).toBe('2026-04-01');
    expect(serialToIso(46160)).toBe('2026-05-18');
    expect(serialToIso(46113.75)).toBe('2026-04-01');
    expect(serialToIso(0)).toBeNull();
  });

  it.each(['dd-mmm-yyyy', 'm/d/yy', 'dd/mm/yyyy hh:mm', '[$-409]d-mmm-yy'])('knows %s is a date format', (f) => expect(isDateNumberFormat(f)).toBe(true));
  it.each(['#,##0.00', 'General', '@', '0%', '"dd" 0.00', undefined])('knows %p is not', (f) => expect(isDateNumberFormat(f)).toBe(false));
});

describe('resolveColumn', () => {
  const t = table();
  it.each([['Debit', 'Debit'], ['debit', 'Debit'], ['ref no', 'Ref No'], ['Ref', 'Ref No'], ['E', 'Debit'], ['value-date', 'Value Date']])(
    'finds %s', (name, expected) => expect(resolveColumn(t, name)?.name).toBe(expected),
  );
  it.each(['Amount', 'Date of birth', '', 'ZZ'])('does not invent a column for %p', (name) => expect(resolveColumn(t, name)).toBeNull());
  it('does not guess between two columns that both match', () => {
    // "Date" matches the Date column exactly, so it is not ambiguous with Value Date.
    expect(resolveColumn(t, 'Date')?.name).toBe('Date');
    const two = buildTable({ sheetName: 'S', sheetData: [['Txn Date', 'Value Date'], [D('2026-04-01'), D('2026-04-01')]] });
    expect(resolveColumn(two, 'date')).toBeNull();
  });
});

describe('executeTableQuery', () => {
  it('sums a column exactly', () => {
    const [debit, credit] = run({ queries: [{ op: 'sum', column: 'Debit' }, { op: 'sum', column: 'Credit' }] });
    expect(debit).toMatchObject({ kind: 'scalar', value: 73098.96, matchedRows: 8 });
    expect(credit).toMatchObject({ kind: 'scalar', value: 50220.78 });
    expect(sumExact([0.1, 0.2])).toBe(0.3);
    expect(sumExact([6173.55, 1925.41, 15000])).toBe(23098.96);
  });

  it('counts rows, and rows that have a value in a column', () => {
    const [all, debits, credits] = run({ queries: [{ op: 'count' }, { op: 'count', column: 'Debit' }, { op: 'count', column: 'Credit' }] });
    expect([all, debits, credits].map((r) => (r.kind === 'scalar' ? r.value : null))).toEqual([8, 6, 2]);
  });

  it('ranks the largest rows, keeps sheet order between equals, and reports the tie it cut', () => {
    const [top] = run({ queries: [{ op: 'top', column: 'Debit', limit: 3 }] });
    if (top.kind !== 'rows') throw new Error('rows expected');
    // 20,000 first, then the 15,000s in sheet order; the third 15,000 is the tie left out.
    expect(top.rows.map((r) => r.rowNumber)).toEqual([9, 5, 6]);
    expect(top.tiedBeyondLimit).toBe(1);
  });

  it('gives the same list every time for the same question', () => {
    const plan = { queries: [{ op: 'top', column: 'Debit', limit: 4 }] };
    const once = JSON.stringify(run(plan));
    for (let i = 0; i < 5; i++) expect(JSON.stringify(run(plan))).toBe(once);
  });

  it('finds the lowest balance together with the row it is on', () => {
    const [min] = run({ queries: [{ op: 'min', column: 'Balance' }] });
    expect(min).toMatchObject({ kind: 'scalar', value: 2121.82, tiedRows: 0 });
    expect(min.kind === 'scalar' && min.row?.rowNumber).toBe(9);
  });

  it('reads the period from the date column', () => {
    const [from, to] = run({ queries: [{ op: 'min', column: 'Date' }, { op: 'max', column: 'Date' }] });
    expect([from, to].map((r) => (r.kind === 'scalar' ? r.value : null))).toEqual(['2026-04-01', '2026-05-18']);
  });

  it('reads the closing balance as the last Balance', () => {
    const [last] = run({ queries: [{ op: 'last', column: 'Balance' }] });
    expect(last).toMatchObject({ kind: 'scalar', value: 2121.82 });
  });

  it('filters by text, ignoring case', () => {
    const [count, total] = run({
      queries: [
        { op: 'count', filters: [{ column: 'Description', op: 'contains', value: 'atm wdl' }] },
        { op: 'sum', column: 'Debit', filters: [{ column: 'Description', op: 'contains', value: 'ATM WDL' }] },
      ],
    });
    expect([count, total].map((r) => (r.kind === 'scalar' ? r.value : null))).toEqual([2, 21173.55]);
  });

  it('filters by a date range and by an amount', () => {
    const [april, large] = run({
      queries: [
        { op: 'sum', column: 'Debit', filters: [{ column: 'Date', op: 'between', value: '2026-04-01', value2: '2026-04-30' }] },
        { op: 'count', filters: [{ column: 'Debit', op: 'gt', value: 10000 }] },
      ],
    });
    expect(april).toMatchObject({ value: 23098.96, matchedRows: 4 });
    expect(large).toMatchObject({ value: 4 });
  });

  it('handles blank and not-blank filters', () => {
    const [credits] = run({ queries: [{ op: 'list', filters: [{ column: 'Debit', op: 'blank' }] }] });
    expect(credits.kind === 'rows' && credits.rows.map((r) => r.rowNumber)).toEqual([2, 8]);
  });

  it('groups by month in date order', () => {
    const [byMonth] = run({ queries: [{ op: 'sum', column: 'Debit', groupBy: { column: 'Date', by: 'month' } }] });
    expect(byMonth.kind === 'groups' && byMonth.groups).toEqual([
      { key: '2026-04', label: 'Apr 2026', value: 23098.96, rows: 4 },
      { key: '2026-05', label: 'May 2026', value: 50000, rows: 4 },
    ]);
  });

  it('keeps the same month of different years apart', () => {
    const t = buildTable({ sheetName: 'S', sheetData: [['Date', 'Amt'], ['2025-04-10', 1], ['2026-04-10', 2]] });
    const [g] = run({ queries: [{ op: 'sum', column: 'Amt', groupBy: { column: 'Date', by: 'month' } }] }, t);
    expect(g.kind === 'groups' && g.groups.map((x) => x.label)).toEqual(['Apr 2025', 'Apr 2026']);
  });

  it('returns nothing, not zero, when no row has a value', () => {
    const [none] = run({ queries: [{ op: 'max', column: 'Debit', filters: [{ column: 'Description', op: 'contains', value: 'nowhere' }] }] });
    expect(none).toMatchObject({ kind: 'scalar', value: null, matchedRows: 0 });
  });

  it('refuses to add up a text column', () => {
    expect(() => run({ queries: [{ op: 'sum', column: 'Description' }] })).toThrow(TableQueryPlanError);
  });
});

describe('parseTableQueryPlan', () => {
  const parse = (plan: unknown) => parseTableQueryPlan(table(), typeof plan === 'string' ? plan : JSON.stringify(plan));

  it('accepts a plan wrapped in prose or a code fence', () => {
    expect(parse('Here you go:\n```json\n{"queries":[{"op":"sum","column":"Debit"}]}\n```').kind).toBe('queries');
  });

  it('accepts the common alternative spellings', () => {
    const plan = parse({ queries: [{ op: 'total', column: 'debit', where: [{ column: 'Debit', operator: '>', value: 5 }], group_by: 'Description' }] });
    expect(plan.kind === 'queries' && plan.queries[0]).toMatchObject({ op: 'sum', filters: [{ op: 'gt', value: 5 }], groupBy: { by: 'value' } });
  });

  it('passes on a refusal', () => {
    expect(parse({ unsupported: 'asks for advice' })).toEqual({ kind: 'unsupported', reason: 'asks for advice' });
  });

  it.each([
    ['a column that is not on the sheet', { queries: [{ op: 'sum', column: 'Amount' }] }],
    ['an unknown operation', { queries: [{ op: 'median', column: 'Debit' }] }],
    ['an unknown filter', { queries: [{ op: 'count', filters: [{ column: 'Debit', op: 'regex', value: 'x' }] }] }],
    ['a filter with no value', { queries: [{ op: 'count', filters: [{ column: 'Debit', op: 'gt' }] }] }],
    ['a between with one end', { queries: [{ op: 'count', filters: [{ column: 'Date', op: 'between', value: '2026-04-01' }] }] }],
    ['a measure with no column', { queries: [{ op: 'top' }] }],
    ['no queries', { queries: [] }],
    ['seven queries', { queries: Array.from({ length: 7 }, () => ({ op: 'count' })) }],
    ['a bad limit', { queries: [{ op: 'top', column: 'Debit', limit: 'many' }] }],
    ['text that is not JSON', 'I think the total is 5000'],
    // TASKS.md #395: "Total debits per month" came back as one line per Debit amount.
    ['a column summed and grouped by itself', { queries: [{ op: 'sum', column: 'Debit', groupBy: { column: 'Debit', by: 'value' } }] }],
  ])('rejects %s', (_label, plan) => {
    expect(() => parse(plan)).toThrow(TableQueryPlanError);
  });

  it('still allows counting how often each value of a column occurs', () => {
    const plan = parse({ queries: [{ op: 'count', column: 'Debit', groupBy: { column: 'Debit', by: 'value' } }] });
    expect(plan.kind).toBe('queries');
  });

  it('accepts a sum of Debit grouped by Date by month', () => {
    const plan = parse({ queries: [{ op: 'sum', column: 'Debit', groupBy: { column: 'Date', by: 'month' } }] });
    expect(plan.kind === 'queries' && plan.queries[0].groupBy).toMatchObject({ by: 'month' });
  });

  it('caps a huge limit', () => {
    const plan = parse({ queries: [{ op: 'top', column: 'Debit', limit: 5000 }] });
    expect(plan.kind === 'queries' && plan.queries[0].limit).toBe(50);
  });
});

describe('formatTableAnswer', () => {
  it('prints numbers and dates itself, never serials', () => {
    expect(formatNumber(281128.12, 2)).toBe('2,81,128.12');
    expect(formatNumber(834, 0)).toBe('834');
    expect(formatDate('2026-07-24')).toBe('24-Jul-2026');
  });

  it('answers totals with what was computed and from how many rows', () => {
    expect(answer({ queries: [{ op: 'sum', column: 'Debit' }, { op: 'sum', column: 'Credit' }] })).toBe(
      // "From 6 rows": the six that have a debit, not all eight.
      '**Total of Debit: 73,098.96**\n   From 6 rows.\n\n**Total of Credit: 50,220.78**\n   From 2 rows.\n\nWorked out from all 8 rows of Bank Statement.',
    );
  });

  it('words a filtered count plainly', () => {
    expect(answer({ queries: [{ op: 'count', column: 'Description', filters: [{ column: 'Description', op: 'starts_with', value: 'CHQ PAID' }] }] })).toContain(
      '**Number of rows where Description starts with "CHQ PAID": 3**',
    );
    expect(answer({ queries: [{ op: 'count', column: 'Debit' }] })).toContain('**Number of rows with a Debit: 6**');
  });

  it('lists the largest debits with date, description, amount and row', () => {
    const text = answer({ queries: [{ op: 'top', column: 'Debit', limit: 3 }] });
    expect(text.split('\n')).toEqual([
      '**3 largest by Debit**',
      '1. 18-May-2026 · CHQ PAID-MICR CTS-FUEL STATION · Debit 20,000.00 (row 9)',
      '2. 28-Apr-2026 · CHQ PAID-MICR CTS-BOOK HOUSE · Debit 15,000.00 (row 5)',
      '3. 06-May-2026 · ATM WDL 605256 SAMPLE TOWN · Debit 15,000.00 (row 6)',
      '1 row more has the same Debit (15,000.00) as the last one listed.',
      '',
      'Worked out from all 8 rows of Bank Statement.',
    ]);
    expect(text).not.toMatch(/\b46\d{3}\b/);
  });

  it('says in words which filter was applied', () => {
    expect(answer({ queries: [{ op: 'sum', column: 'Debit', filters: [{ column: 'Description', op: 'contains', value: 'ATM WDL' }] }] })).toContain(
      '**Total of Debit where Description contains "ATM WDL": 21,173.55**',
    );
    expect(answer({ queries: [{ op: 'sum', column: 'Debit', filters: [{ column: 'Date', op: 'between', value: '2026-04-01', value2: '2026-04-30' }] }] })).toContain(
      '**Total of Debit where Date is between 01-Apr-2026 and 30-Apr-2026: 23,098.96**',
    );
  });

  it('gives the date and row of a lowest value', () => {
    expect(answer({ queries: [{ op: 'min', column: 'Balance' }] })).toContain(
      '**Lowest Balance: 2,121.82**\n   18-May-2026 · CHQ PAID-MICR CTS-FUEL STATION (row 9).',
    );
  });

  it('gives a period as two dates', () => {
    const text = answer({ queries: [{ op: 'min', column: 'Date' }, { op: 'max', column: 'Date' }] });
    // Just the two dates: no row description under either.
    expect(text).toBe('**Earliest Date: 01-Apr-2026**\n\n**Latest Date: 18-May-2026**\n\nWorked out from all 8 rows of Bank Statement.');
  });

  it('says so when nothing matches, and does not print a zero', () => {
    expect(answer({ queries: [{ op: 'sum', column: 'Debit', filters: [{ column: 'Description', op: 'contains', value: 'nowhere' }] }] })).toContain(
      'nothing to work from. No rows match.',
    );
  });

  it('says when not every row of the sheet could be read', () => {
    const t = table();
    const [r] = run({ queries: [{ op: 'count' }] }, t);
    expect(formatTableAnswer(t, [r], { expectedRows: 834 })).toMatch(/Only 8 rows of 834 on Bank Statement could be read, so this may be incomplete\.$/);
  });
});

describe('looksLikeTableQuestion', () => {
  it.each([
    'Which are the 10 largest debits in Bank Statement?',
    'What is the total of Debit and the total of Credit in Bank Statement?',
    'How many ATM withdrawals are there, and what do they add up to?',
    'What period does this statement cover?',
    'What was the lowest balance and on which date?',
    'How much did I spend at FUEL STATION?',
    'List the transactions above 10000',
    'show me the 5 biggest deposits',
    // TASKS.md #394: a phrase that opens with the measure is a question too.
    'Total debits to samplepayee000',
    'Total debits per month',
    'Sum of debits',
    'Count of ATM withdrawals',
    'Top 5 debits',
    'Average balance',
    'Number of cheque transactions',
  ])('takes "%s"', (m) => expect(looksLikeTableQuestion(m)).toBe(true));

  it.each([
    'Highlight every row where Debit is more than 10000',
    'Sort Bank Statement by Debit, largest first',
    'Add a total row at the bottom',
    'Total the Debit column in a new row',
    'Sum column E into cell F1',
    'Sort the top 10 debits to the top',
    'Highlight the largest debit',
    'Can you create a chart of the balance?',
    'What is in cell E3?',
    'How many sheets are in this workbook?',
    'Explain what this workbook is for',
    'hi',
    '',
  ])('leaves "%s" alone', (m) => expect(looksLikeTableQuestion(m)).toBe(false));
});

describe('commonStarts', () => {
  const description = { index: 2 };

  it('lists the wordings a description column uses, most common first', () => {
    expect(commonStarts(table(), description)).toEqual([
      { text: 'CHQ PAID', rows: 3 },
      { text: 'ATM WDL', rows: 2 },
      { text: 'INT.PD', rows: 1 },
      { text: 'NEFT CR', rows: 1 },
      { text: 'UPI', rows: 1 },
    ]);
  });

  it('shows them to the planner, so "interest" can be matched to "INT.PD"', () => {
    const message = buildTableQueryUserMessage(table(), 'How much interest was credited?');
    expect(message).toContain('entries begin with: "CHQ PAID" x3, "ATM WDL" x2, "INT.PD" x1, "NEFT CR" x1, "UPI" x1');
    // Only the Description column has a vocabulary worth listing.
    expect(message.match(/entries begin with/g)).toHaveLength(1);
  });

  it('lists nothing for a column of free text, where every entry starts differently', () => {
    const names = Array.from({ length: 20 }, (_, i) => String.fromCharCode(65 + i).repeat(3) + ' Person');
    const t = buildTable({ sheetName: 'S', sheetData: [['Name'], ...names.map((n) => [n])] });
    expect(commonStarts(t, { index: 0 })).toEqual([]);
  });

  it('stops at the first digit or separator and keeps at most three words', () => {
    const t = buildTable({
      sheetName: 'S',
      sheetData: [['D'], ...Array.from({ length: 4 }, (_, i) => [`SMS ALERT CHARGES INCL GST ${i}`]), ...Array.from({ length: 4 }, (_, i) => [`UPIOUT/12345${i}/payee@upi`])],
    });
    expect(commonStarts(t, { index: 0 })).toEqual([{ text: 'SMS ALERT CHARGES', rows: 4 }, { text: 'UPIOUT', rows: 4 }]);
  });
});

describe('TableQuestionAnswerer', () => {
  const input = { message: 'Which are the 2 largest debits?', sheetName: 'Bank Statement', sheetData: [HEADERS, ...ROWS], columnHints: HINTS };
  const withReply = (reply: string | Error, truncated = false) => {
    const complete = jest.fn(async (opts: { outcome?: { truncated?: boolean } }) => {
      if (reply instanceof Error) throw reply;
      if (opts.outcome) opts.outcome.truncated = truncated;
      return reply;
    });
    return { answerer: new TableQuestionAnswerer({ complete } as never), complete };
  };

  it('asks the model for a plan only: no row of data is sent beyond a few sample cells', async () => {
    const { answerer, complete } = withReply('{"queries":[{"op":"top","column":"Debit","limit":2}]}');
    await answerer.tryAnswer(input);
    const sent = complete.mock.calls[0][0] as unknown as { userMessage: string; responseFormat: string };
    expect(sent.responseFormat).toBe('json_object');
    expect(sent.userMessage).toBe(buildTableQueryUserMessage(table(), input.message));
    // Column names, types and three samples each. Not the 20,000 debit on row 9.
    expect(sent.userMessage).toContain('- "Date" (date) e.g. "2026-04-01", "2026-04-03", "2026-04-06"');
    expect(sent.userMessage).not.toContain('20000');
    expect(sent.userMessage).not.toContain('FUEL STATION');
  });

  it('computes and words the answer itself', async () => {
    const { answerer } = withReply('{"queries":[{"op":"top","column":"Debit","limit":2}]}');
    const result = await answerer.tryAnswer(input);
    expect(result?.answer).toBe(
      '**2 largest by Debit**\n' +
        '1. 18-May-2026 · CHQ PAID-MICR CTS-FUEL STATION · Debit 20,000.00 (row 9)\n' +
        '2. 28-Apr-2026 · CHQ PAID-MICR CTS-BOOK HOUSE · Debit 15,000.00 (row 5)\n' +
        '2 rows more have the same Debit (15,000.00) as the last one listed.\n\n' +
        'Worked out from all 8 rows of Bank Statement.',
    );
    // The rows it names, in the order it names them, for the add-in to point at.
    expect(result?.rowNumbers).toEqual([9, 5]);
    expect(result?.columnCount).toBe(9);
  });

  it('names exactly the rows its text names', async () => {
    const plan = { queries: [{ op: 'min', column: 'Balance' }, { op: 'top', column: 'Debit', limit: 2 }, { op: 'min', column: 'Date' }, { op: 'sum', column: 'Debit' }] };
    const { answerer } = withReply(JSON.stringify(plan));
    const result = await answerer.tryAnswer(input);
    const inText = [...(result?.answer ?? '').matchAll(/\(row (\d+)\)/g)].map((m) => Number(m[1]));
    // Row 9 is both the lowest balance and the largest debit: named twice, pointed at once.
    expect(inText).toEqual([9, 9, 5]);
    expect(result?.rowNumbers).toEqual([9, 5]);
  });

  it.each([
    ['the planner declines', '{"unsupported":"asks for advice"}', false],
    ['the plan names a column that does not exist', '{"queries":[{"op":"sum","column":"Amount"}]}', false],
    ['the reply is not JSON', 'The largest is 20,000', false],
    ['the reply was cut off', '{"queries":[{"op":"sum","column":"Debit"}]}', true],
    ['the model call fails', new Error('timeout'), false],
  ])('steps aside when %s', async (_label, reply, truncated) => {
    const { answerer } = withReply(reply as string | Error, truncated as boolean);
    expect(await answerer.tryAnswer(input)).toBeNull();
  });

  it('does not call the model for an empty sheet', async () => {
    const { answerer, complete } = withReply('{}');
    expect(await answerer.tryAnswer({ ...input, sheetData: [HEADERS] })).toBeNull();
    expect(complete).not.toHaveBeenCalled();
  });
});
