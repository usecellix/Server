/**
 * Stage 2 deterministic lanes: P2.2 fuzzy match, P2.5 days outstanding (+ P3.11 / P3.14
 * formula preview & verification), P2.6 PAN, P2.7 fixed-asset register, P2.8 narrations.
 */
import { BatchSetOperation, SheetActionPayload } from '../types/sheet-actions.types';
import { PAN_ENTITY_TYPES, PAN_REGEX } from '../utils/india-validators.util';
import { parseNarration } from '../../domain-tools/bank/narration-parser';
import { getRuleTable } from '../../rules/rule-tables';
import { DEFAULT_BANDS, matchOneToOne, similarity } from '../../matching/matching-engine';
import { DeterministicLane, LaneContext, LaneResult } from './lane.types';
import {
  columnLetter,
  columnsFromMessage,
  dateInMessage,
  daysBetween,
  excelDateLiteral,
  findHeader,
  formatDdMmYyyy,
  headerNamedInMessage,
  lastUsedColumn,
  lastUsedRow,
  parseDateValue,
  str,
  targetCellFromMessage,
  toNumber,
} from './lane-helpers';

const header = (ctx: LaneContext): unknown[] => ctx.values[ctx.headerRowIndex] ?? [];
const dataRows = (ctx: LaneContext): number[] => {
  const last = lastUsedRow(ctx.values);
  const rows: number[] = [];
  for (let r = ctx.headerRowIndex + 1; r <= last; r += 1) rows.push(r);
  return rows;
};

/** Reuse a column we wrote before (same header) or take the next free one. */
function outputColumns(ctx: LaneContext, names: string[]): number[] {
  const h = header(ctx);
  let next = lastUsedColumn(ctx.values) + 1;
  return names.map((name) => {
    const existing = h.findIndex((x) => str(x).trim() === name);
    return existing >= 0 ? existing : next++;
  });
}

const addr = (col: number, row: number): string => `${columnLetter(col)}${row + 1}`;

// ---------------------------------------------------------------- P2.2 fuzzy match

function sheetsInMessage(message: string, names: string[]): string[] {
  return names
    .map((n) => {
      const esc = n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const m = new RegExp(`(^|[^A-Za-z0-9_])${esc}($|[^A-Za-z0-9_])`, 'i').exec(message);
      return { n, i: m ? m.index : -1 };
    })
    .filter((x) => x.i >= 0)
    .sort((a, b) => a.i - b.i)
    .map((x) => x.n);
}

export const fuzzyMatchLane: DeterministicLane = {
  step: 'P2.2',
  id: 'fuzzy-match',
  writes: true,
  detect: (m) =>
    /\b(match|fuzzy|link)\b/i.test(m) &&
    /\b(vendor|party|parties|supplier|customer|name|names|ledger names?)\b/i.test(m) &&
    !/\b(gstr|2a|2b|reconcil\w*|bank)\b/i.test(m),
  run: async (ctx) => {
    const named = sheetsInMessage(ctx.message, ctx.sheetNames);
    const leftSheet = named.length >= 2 ? named[0] : ctx.sheetName;
    const rightSheet = named.length >= 2 ? named[1] : named.find((n) => n !== ctx.sheetName);
    if (!rightSheet) {
      return { kind: 'answer', answer: 'Tell me which two sheets to match, e.g. *Match the vendor names in T_FuzzyA column A with T_FuzzyB column A*.' };
    }
    const left = leftSheet === ctx.sheetName ? ctx.values : await ctx.getSheetValues(leftSheet);
    const right = await ctx.getSheetValues(rightSheet);
    if (!left || !right) return { kind: 'answer', answer: `I could not read ${!left ? leftSheet : rightSheet}.` };

    const cols = columnsFromMessage(ctx.message);
    const pick = (vals: unknown[][], given?: number) =>
      given ?? Math.max(0, findHeader(vals[0] ?? [], /vendor|party|supplier|customer|name/i));
    const lc = pick(left, cols[0]);
    const rc = pick(right, cols[1] ?? cols[0]);
    const lRows = left.slice(1).map((r, i) => ({ row: i + 1, name: str(r?.[lc]).trim() })).filter((x) => x.name);
    const rRows = right.slice(1).map((r) => str(r?.[rc]).trim()).filter(Boolean);

    const bands = { ...DEFAULT_BANDS };
    const result = matchOneToOne(lRows, rRows, (l, r) => similarity(l.name, r), bands);
    const includeReview = /\b(review|all bands?|include (the )?review|also link)\b/i.test(ctx.message);

    const groups = { high: [] as string[], review: [] as string[], none: [] as string[] };
    for (const c of result) {
      const line = c.right ? `${c.left.name} ↔ ${c.right} (${c.score})` : `${c.left.name} — no match`;
      groups[c.band].push(line);
    }
    const toLink = result.filter((c) => c.band === 'high' || (includeReview && c.band === 'review'));

    const lines = [
      `Matched **${lRows.length}** names in ${leftSheet} column ${columnLetter(lc)} against ${rightSheet} column ${columnLetter(rc)}.`,
      '',
      `**High confidence (≥${bands.high})** — ${groups.high.length}`,
      ...groups.high.map((l) => `- ${l}`),
      `**Review (${bands.review}–${bands.high - 1})** — ${groups.review.length}`,
      ...groups.review.map((l) => `- ${l}`),
      `**No match** — ${groups.none.length}`,
      ...groups.none.map((l) => `- ${l}`),
      '',
      includeReview
        ? 'Accept links the high **and** review pairs.'
        : 'Accept links the **high** band only. To link the review band too, ask: *also link the review band*.',
    ];
    if (toLink.length === 0) return { kind: 'answer', answer: lines.join('\n') };

    const outCtx: LaneContext = { ...ctx, values: left, headerRowIndex: 0 };
    const [mCol, sCol, bCol] = outputColumns(outCtx, [`Matched in ${rightSheet}`, 'Match Score', 'Match Band']);
    const ops: BatchSetOperation[] = [
      { address: addr(mCol, 0), value: `Matched in ${rightSheet}` },
      { address: addr(sCol, 0), value: 'Match Score' },
      { address: addr(bCol, 0), value: 'Match Band' },
    ];
    for (const c of toLink) {
      ops.push({ address: addr(mCol, c.left.row), value: c.right });
      ops.push({ address: addr(sCol, c.left.row), value: c.score });
      ops.push({ address: addr(bCol, c.left.row), value: c.band === 'high' ? 'High' : 'Review' });
    }
    return {
      kind: 'actions',
      answer: lines.join('\n'),
      explanation: `Link ${toLink.length} matched names.`,
      actions: [{ type: 'BATCH_SET', sheetName: leftSheet, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P2.5 days outstanding (+ P3.11, P3.14)

function dateColumn(ctx: LaneContext): number {
  const h = header(ctx);
  const named = headerNamedInMessage(h, ctx.message);
  if (named >= 0 && /date/i.test(str(h[named]))) return named;
  const byHeader = findHeader(h, /invoice date|bill date|^date$|date/i);
  if (byHeader >= 0) return byHeader;
  return -1;
}

export interface DaysRow {
  row: number;
  raw: unknown;
  days: number | null;
  overdue: number | null;
  note: string;
}

export function computeDays(ctx: LaneContext, dateCol: number, asOn: number, creditCol: number): DaysRow[] {
  return dataRows(ctx).map((r) => {
    const raw = ctx.values[r]?.[dateCol];
    const ms = parseDateValue(raw);
    if (ms === null) {
      const note = str(raw).trim() === '' ? 'blank date → left blank' : `"${str(raw)}" is not a date → left blank`;
      return { row: r, raw, days: null, overdue: null, note };
    }
    const days = daysBetween(ms, asOn);
    const credit = creditCol >= 0 ? toNumber(ctx.values[r]?.[creditCol]) ?? 0 : 0;
    return { row: r, raw, days, overdue: creditCol >= 0 ? days - credit : null, note: days < 0 ? 'future-dated' : '' };
  });
}

export const daysOutstandingLane: DeterministicLane = {
  step: 'P2.5',
  id: 'days-outstanding',
  writes: true,
  detect: (m) =>
    /\b(days?\s+(outstanding|overdue|between|since|old)|overdue\s+days|days\s+past\s+due|age\s+in\s+days)\b/i.test(m),
  run: (ctx) => {
    const asOn = dateInMessage(ctx.message);
    if (asOn === null) {
      return {
        kind: 'answer',
        answer:
          'As on which date should I count the days? Please give the date, e.g. *…as on 31-03-2026*. I will not assume today’s date silently. (Say *as on today* to use a live TODAY() formula.)',
      };
    }
    const h = header(ctx);
    const dCol = dateColumn(ctx);
    if (dCol < 0) return { kind: 'answer', answer: 'I could not find a date column. Name it, e.g. *from Invoice Date*.' };
    const wantOverdue = /\b(overdue|beyond credit|past due|credit days)\b/i.test(ctx.message);
    const creditCol = wantOverdue ? findHeader(h, /credit\s*(days|period)/i) : -1;
    const verify = /\bverif(y|ication)\b|\btest (it|the formula)\b|\bedge cases?\b/i.test(ctx.message);
    const target = targetCellFromMessage(ctx.message);

    const rows = computeDays(ctx, dCol, asOn, creditCol);
    const first = rows[0]?.row ?? ctx.headerRowIndex + 1;
    const dateLit = excelDateLiteral(asOn);
    const d = (r: number) => addr(dCol, r);

    const names = ['Days Outstanding', ...(creditCol >= 0 ? ['Overdue Days'] : []), 'Date Check'];
    const cols = target
      ? names.map((_, i) => target.col + i)
      : outputColumns(ctx, names);
    const ops: BatchSetOperation[] = [];
    names.forEach((n, i) => ops.push({ address: addr(cols[i], ctx.headerRowIndex), value: n }));
    for (const row of rows) {
      const r = row.row;
      ops.push({ address: addr(cols[0], r), formula: `=IF(ISNUMBER(${d(r)}),${dateLit}-${d(r)},"")` });
      if (creditCol >= 0) {
        ops.push({ address: addr(cols[1], r), formula: `=IF(ISNUMBER(${d(r)}),${dateLit}-${d(r)}-N(${addr(creditCol, r)}),"")` });
      }
      ops.push({
        address: addr(cols[names.length - 1], r),
        formula: `=IF(NOT(ISNUMBER(${d(r)})),IF(${d(r)}="","Blank date","Not a date"),IF(${d(r)}>${dateLit},"Future-dated",""))`,
      });
    }

    const samples = rows.slice(0, 5).map((x) => (x.days === null ? '—' : String(x.days)));
    const overdueSamples = rows.slice(0, 5).map((x) => (x.overdue === null ? '—' : String(x.overdue)));
    const future = rows.filter((x) => x.note === 'future-dated');
    const lines = [
      `**Plain English:** for each row, days from **${str(h[dCol]) || columnLetter(dCol)}** (column ${columnLetter(dCol)}) to **${formatDdMmYyyy(asOn)}**${creditCol >= 0 ? `, minus **${str(h[creditCol])}**` : ''}. Live formulas — they recalculate if a date changes.`,
      `**First results:** ${samples.join(', ')}${creditCol >= 0 ? `  ·  overdue beyond credit days: ${overdueSamples.join(', ')}` : ''}`,
      `Formula in ${addr(cols[0], first)}: \`=IF(ISNUMBER(${d(first)}),${dateLit}-${d(first)},"")\``,
    ];
    if (future.length) {
      lines.push(`⚠ ${future.length} future-dated row(s) — shown as negative and flagged “Future-dated”, not hidden as 0: rows ${future.map((f) => f.row + 1).join(', ')}.`);
    }
    if (verify) {
      lines.push('', '**Verification on edge cases:**');
      const edge = rows.filter((x) => x.days === null || x.days <= 0);
      if (!edge.length) lines.push('- No blank, zero, negative or text dates in this data.');
      for (const x of edge) {
        lines.push(`- Row ${x.row + 1}: ${x.days === null ? x.note : x.days === 0 ? '0 (same day)' : `${x.days} (future-dated, flagged)`}`);
      }
      lines.push('- Text or blank dates return blank instead of #VALUE! (guarded with ISNUMBER).');
    }
    return {
      kind: 'actions',
      answer: lines.join('\n'),
      explanation: `Add ${names.join(' / ')} as on ${formatDdMmYyyy(asOn)}.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P2.6 PAN

export type PanCheck = { pan: string; status: string; entity: string };

export function checkPanCell(raw: unknown): PanCheck {
  const text = str(raw).trim();
  if (!text) return { pan: '', status: '', entity: '' };
  const single = /^\S+$/.test(text);
  if (single) {
    if (PAN_REGEX.test(text)) {
      const entity = PAN_ENTITY_TYPES[text[3]];
      return entity ? { pan: text, status: 'Valid', entity } : { pan: text, status: 'Invalid entity type', entity: '' };
    }
    if (PAN_REGEX.test(text.toUpperCase())) {
      return { pan: text, status: 'Check — lowercase (not auto-corrected)', entity: PAN_ENTITY_TYPES[text[3].toUpperCase()] ?? '' };
    }
    return { pan: text, status: 'Invalid format', entity: '' };
  }
  const found = text.match(/\b[A-Z]{5}[0-9]{4}[A-Z]\b/);
  if (!found) return { pan: '', status: 'No PAN found', entity: '' };
  const pan = found[0];
  const entity = PAN_ENTITY_TYPES[pan[3]];
  return { pan, status: entity ? 'Valid (extracted)' : 'Invalid entity type', entity: entity ?? '' };
}

export const panLane: DeterministicLane = {
  step: 'P2.6',
  id: 'pan',
  writes: true,
  detect: (m) => /\bpans?\b/i.test(m) && /\b(validat\w*|check|verify|extract|entity)\b/i.test(m) && !/\bgstins?\b/i.test(m),
  run: (ctx) => {
    const h = header(ctx);
    const col = columnsFromMessage(ctx.message)[0] ?? Math.max(0, findHeader(h, /\bpan\b|text|narration|remark/i));
    const [pCol, sCol, eCol] = outputColumns(ctx, ['PAN', 'PAN Status', 'Entity Type']);
    const ops: BatchSetOperation[] = [
      { address: addr(pCol, ctx.headerRowIndex), value: 'PAN' },
      { address: addr(sCol, ctx.headerRowIndex), value: 'PAN Status' },
      { address: addr(eCol, ctx.headerRowIndex), value: 'Entity Type' },
    ];
    const counts: Record<string, number> = {};
    const review: string[] = [];
    for (const r of dataRows(ctx)) {
      const res = checkPanCell(ctx.values[r]?.[col]);
      if (!res.status) continue;
      counts[res.status] = (counts[res.status] ?? 0) + 1;
      ops.push({ address: addr(pCol, r), value: res.pan });
      ops.push({ address: addr(sCol, r), value: res.status });
      ops.push({ address: addr(eCol, r), value: res.entity });
      if (!res.status.startsWith('Valid')) review.push(`row ${r + 1} — ${str(ctx.values[r]?.[col])} → ${res.status}`);
    }
    const summary = Object.entries(counts).map(([k, v]) => `${v} ${k}`).join(', ');
    return {
      kind: 'actions',
      answer: [
        `Checked PANs in column ${columnLetter(col)}: ${summary}. I'll add **PAN**, **PAN Status** and **Entity Type** (4th character: P Individual, C Company, H HUF, F Firm/LLP, T Trust…). Nothing is auto-corrected.`,
        ...(review.length ? ['', 'For review:', ...review.map((x) => `- ${x}`)] : []),
      ].join('\n'),
      explanation: 'Validate PANs and add entity type.',
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P2.7 fixed asset register

/** Frozen FAR layout — the colleague's depreciation (C7.4/C7.5) reads these exact headers. */
export const FAR_LAYOUT = [
  'Asset',
  'Block/Class',
  'Date of Purchase',
  'Put-to-use Date',
  'Cost',
  'Additions',
  'Deletions',
  'Location',
  'Tag No.',
] as const;
export const FAR_SHEET_NAME = 'Fixed Asset Register';
export const INDIAN_NUMBER_FORMAT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';

export const farTemplateLane: DeterministicLane = {
  step: 'P2.7',
  id: 'far-template',
  writes: true,
  detect: (m) => /\b(fixed asset register|asset register|\bFAR\b)\b/i.test(m) && /\b(create|make|build|set up|setup|template|new|generate)\b/i.test(m),
  run: () => {
    const name = FAR_SHEET_NAME;
    const actions: SheetActionPayload[] = [
      { type: 'CREATE_SHEET', sheetName: name, name, uniqueName: true },
      { type: 'WRITE_TABLE', sheetName: name, headers: [...FAR_LAYOUT], rows: [] },
      { type: 'FORMAT_RANGE', sheetName: name, range: 'A1:I1', format: { bold: true, fillColor: '#203764', fontColor: '#FFFFFF' } },
      { type: 'FORMAT_RANGE', sheetName: name, range: 'C2:D1000', format: { numberFormat: 'dd-mm-yyyy' } },
      { type: 'FORMAT_RANGE', sheetName: name, range: 'E2:G1000', format: { numberFormat: INDIAN_NUMBER_FORMAT } },
      { type: 'FREEZE_PANES', sheetName: name, freezeRows: 1 },
      { type: 'AUTOFIT_COLUMNS', sheetName: name },
    ] as SheetActionPayload[];
    return {
      kind: 'actions',
      answer: [
        `I'll create **${name}** with the standard layout: ${FAR_LAYOUT.join(' · ')}.`,
        'Dates as dd-mm-yyyy, Cost/Additions/Deletions in Indian lakh/crore format, header frozen. If the name is taken, a "(2)" copy is created — nothing is overwritten.',
      ].join('\n'),
      explanation: 'Create a fixed asset register.',
      actions,
    };
  },
};

// ---------------------------------------------------------------- P2.8 bank narrations

const IFSC = /^[A-Z]{4}0[A-Z0-9]{6}$/;
const NOISE = /\b(UPI|NEFT|RTGS|IMPS|CHQ|CHEQUE|DEP|DEPOSIT|PAYMENT|TO|FROM|BY|TRF|TRANSFER|P2A|P2M|INV|INVOICE|NO|CLG)\b/gi;

export interface NarrationRow {
  mode: string;
  reference: string;
  counterparty: string;
  invoice: string;
  parsed: boolean;
}

export function parseNarrationRow(raw: unknown): NarrationRow {
  const text = str(raw).trim();
  const p = parseNarration(text);
  const upper = text.toUpperCase();
  let reference = p.referenceNo ?? '';
  if (!reference || IFSC.test(reference)) {
    const tokens = upper.split(/[^A-Z0-9]+/).filter((t) => /\d/.test(t) && t.length >= 12 && !IFSC.test(t));
    reference = tokens.sort((a, b) => b.length - a.length)[0] ?? reference;
  }
  let instrument = '';
  const chq = /\b(?:CHQ|CHEQUE)\b\D{0,12}(\d{3,8})\b/.exec(upper);
  if (chq) instrument = chq[1];
  let counterparty = p.counterparty ?? '';
  counterparty = counterparty.replace(NOISE, ' ').replace(/\s+/g, ' ').trim();
  if (!counterparty && (p.channel || instrument)) {
    counterparty = upper
      .replace(/[^A-Z& ]+/g, ' ')
      .replace(NOISE, ' ')
      .replace(/\b(ATM|WDL|POS|CASH)\b/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }
  // P4.8: per-bank narration packs come from the rule table (B5 stand-in), not inline code.
  const packMode = getRuleTable('bank-narration-packs')
    .data.flatMap((pack) => pack.modes)
    .find(([re]) => re.test(text))?.[1];
  const mode = p.channel ?? packMode ?? (instrument ? 'CHQ' : '');
  const parsed = Boolean(mode);
  return {
    mode,
    reference: mode === 'CHQ' ? instrument || reference : reference,
    counterparty: parsed ? counterparty : '',
    invoice: p.invoiceNumbers.map((x) => x.replace(/[-/]+$/, '')).join(', '),
    parsed,
  };
}

export const narrationLane: DeterministicLane = {
  step: 'P2.8',
  id: 'narration',
  writes: true,
  detect: (m) => /\bnarrations?\b/i.test(m) && /\b(parse|split|extract|break|read|identify)\b/i.test(m),
  run: (ctx) => {
    const h = header(ctx);
    const col = columnsFromMessage(ctx.message)[0] ?? Math.max(0, findHeader(h, /narration|description|particulars|remarks/i));
    const names = ['Mode', 'Reference', 'Counterparty', 'Invoice No'];
    const cols = outputColumns(ctx, names);
    const ops: BatchSetOperation[] = names.map((n, i) => ({ address: addr(cols[i], ctx.headerRowIndex), value: n }));
    const unparsed: string[] = [];
    let parsedCount = 0;
    for (const r of dataRows(ctx)) {
      const raw = ctx.values[r]?.[col];
      if (str(raw).trim() === '') continue;
      const p = parseNarrationRow(raw);
      if (!p.parsed) {
        unparsed.push(`row ${r + 1} — ${str(raw)}`);
        ops.push({ address: addr(cols[0], r), value: 'Unparsed' });
        continue;
      }
      parsedCount += 1;
      [p.mode, p.reference, p.counterparty, p.invoice].forEach((v, i) => ops.push({ address: addr(cols[i], r), value: v }));
    }
    return {
      kind: 'actions',
      answer: [
        `Parsed **${parsedCount}** narrations in column ${columnLetter(col)} into Mode / Reference / Counterparty / Invoice No.`,
        unparsed.length ? `**${unparsed.length} not recognised** (marked "Unparsed", nothing guessed):` : 'Every narration was recognised.',
        ...unparsed.map((u) => `- ${u}`),
      ].join('\n'),
      explanation: 'Parse bank narrations.',
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

export const STAGE2_LANES: DeterministicLane[] = [
  fuzzyMatchLane,
  daysOutstandingLane,
  panLane,
  farTemplateLane,
  narrationLane,
];

export type { LaneResult };
