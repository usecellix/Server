/**
 * Stage 4 lanes that were waiting on the colleague's B5 rule registry. They read every
 * statutory value from `Server/src/rules/rule-tables.ts` (local B5 stand-in) and cite it.
 *   P4.1 compliance calendar · P4.2 validation presets
 */
import { SheetActionPayload } from '../types/sheet-actions.types';
import { CalendarRule, getRuleTable, ruleCitation } from '../../rules/rule-tables';
import { DeterministicLane } from './lane.types';
import { columnLetter, columnsFromMessage, lastUsedRow, toExcelSerial } from './lane-helpers';
import { tableWithFormulas } from './stage3.lanes';

const MONTH = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export type EntityFlag = 'gst-monthly' | 'gst-qrmp' | 'tds' | 'company' | 'audit' | 'msme' | 'any';

export function entityFlags(message: string): Set<EntityFlag> {
  const flags = new Set<EntityFlag>(['any']);
  if (/\bqrmp\b|\bquarterly\s+gst\b/i.test(message)) flags.add('gst-qrmp');
  else if (/\bgst\b/i.test(message)) flags.add('gst-monthly');
  if (/\btds\b/i.test(message)) flags.add('tds');
  if (/\b(company|pvt|private limited|ltd)\b/i.test(message)) flags.add('company');
  if (/\b(tax\s+)?audit\b/i.test(message)) flags.add('audit');
  if (/\bmsme\b/i.test(message)) flags.add('msme');
  return flags;
}

export function fyStartYear(message: string, now: Date): number {
  const m = /\bFY\s*(\d{4})\s*[-–/]\s*(\d{2,4})\b/i.exec(message);
  if (m) return Number(m[1]);
  return now.getUTCMonth() >= 3 ? now.getUTCFullYear() : now.getUTCFullYear() - 1;
}

export interface CalendarRow {
  ret: string;
  period: string;
  due: number; // UTC ms
  section: string;
}

const utc = (y: number, m: number, d: number) => Date.UTC(y, m - 1, d);

export function buildCalendar(rules: CalendarRule[], flags: Set<EntityFlag>, fy: number): CalendarRow[] {
  const rows: CalendarRow[] = [];
  const applies = (r: CalendarRule) => r.needs.some((n) => flags.has(n as EntityFlag));
  for (const rule of rules.filter(applies)) {
    if (rule.frequency === 'monthly') {
      for (let i = 0; i < 12; i += 1) {
        const month = ((3 + i) % 12) + 1; // Apr..Mar
        const year = month >= 4 ? fy : fy + 1;
        const ex = rule.exceptions?.[month];
        const due = ex ? utc(fy + ex[2], ex[0], ex[1]) : utc(month === 12 ? year + 1 : year, (month % 12) + 1, rule.dueDay ?? 1);
        rows.push({ ret: rule.return, period: `${MONTH[month - 1]}-${year}`, due, section: rule.section });
      }
    } else if (rule.frequency === 'quarterly') {
      const ends: Array<[number, number, string]> = [[6, fy, 'Q1'], [9, fy, 'Q2'], [12, fy, 'Q3'], [3, fy + 1, 'Q4']];
      for (const [m, y, q] of ends) {
        rows.push({ ret: rule.return, period: `${q} FY ${fy}-${String(fy + 1).slice(2)}`, due: utc(m === 12 ? y + 1 : y, (m % 12) + 1, rule.dueDay ?? 1), section: rule.section });
      }
    } else {
      for (const [m, d, off, label] of rule.fixed ?? []) {
        rows.push({ ret: rule.return, period: `${label} · FY ${fy}-${String(fy + 1).slice(2)}`, due: utc(fy + off, m, d), section: rule.section });
      }
    }
  }
  return rows.sort((a, b) => a.due - b.due || a.ret.localeCompare(b.ret));
}

export const complianceCalendarLane: DeterministicLane = {
  step: 'P4.1',
  id: 'compliance-calendar',
  writes: true,
  detect: (m) => /\b(compliance|due[-\s]?date)\s+(calendar|schedule)\b/i.test(m) && /\b(create|make|build|generate|set up|setup|new)\b/i.test(m),
  run: (ctx) => {
    const table = getRuleTable('compliance-calendar');
    const flags = entityFlags(ctx.message);
    const fy = fyStartYear(ctx.message, ctx.now);
    const rows = buildCalendar(table.data, flags, fy);
    const client = /\bfor\s+(?:client\s+)?([A-Z][A-Za-z0-9&.' ]+?)(?:\s+with\b|,|$)/.exec(ctx.message)?.[1]?.trim() ?? 'My Client';
    const name = `Compliance Calendar FY ${fy}-${String(fy + 1).slice(2)}`;
    const headers = ['Client', 'Return', 'Period', 'Due date', 'Status', 'Filed on', 'ARN', 'Remarks', 'Rule source'];
    const cite = `${table.id} v${table.version}${table.caVerified ? '' : ' (CA to verify)'}`;
    const body = rows.map((r) => [client, r.ret, r.period, toExcelSerial(r.due), 'Pending', '', '', r.section, cite]);
    const last = rows.length + 1;
    const actions: SheetActionPayload[] = [
      { type: 'CREATE_SHEET', sheetName: name, name, uniqueName: true } as SheetActionPayload,
      ...tableWithFormulas(name, headers, body),
      { type: 'FORMAT_RANGE', sheetName: name, range: 'A1:I1', format: { bold: true, fillColor: '#203764', fontColor: '#FFFFFF' } },
      { type: 'FORMAT_RANGE', sheetName: name, range: `D2:D${last}`, format: { numberFormat: 'dd-mm-yyyy' } },
      { type: 'FORMAT_RANGE', sheetName: name, range: `F2:F${last}`, format: { numberFormat: 'dd-mm-yyyy' } },
      {
        type: 'DATA_VALIDATION',
        sheetName: name,
        range: `E2:E${last}`,
        validation: { kind: 'list', listSource: ['Pending', 'Filed', 'Not applicable'], errorTitle: 'Invalid status', errorMessage: 'Pick Pending, Filed or Not applicable.' },
      },
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: name,
        range: `A2:I${last}`,
        rule: { kind: 'formula', formula: '=AND($E2<>"Filed",$E2<>"Not applicable",$D2<TODAY())', format: { fillColor: '#FFC7CE' } },
      },
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: name,
        range: `A2:I${last}`,
        rule: { kind: 'formula', formula: '=AND($E2<>"Filed",$E2<>"Not applicable",$D2>=TODAY(),$D2-TODAY()<=7)', format: { fillColor: '#FFEB9C' } },
      },
      { type: 'FREEZE_PANES', sheetName: name, freezeRows: 1 },
      { type: 'AUTOFIT_COLUMNS', sheetName: name },
    ] as SheetActionPayload[];
    const by = new Map<string, number>();
    for (const r of rows) by.set(r.ret, (by.get(r.ret) ?? 0) + 1);
    const fmt = (ms: number) => {
      const d = new Date(ms);
      return `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`;
    };
    const firstOf = (ret: string) => rows.find((r) => r.ret === ret);
    return {
      kind: 'actions',
      answer: [
        `**${name}** for ${client} — ${rows.length} rows (${[...by.entries()].map(([k, v]) => `${k} ×${v}`).join(', ')}).`,
        ...['GSTR-1', 'GSTR-3B', 'TDS payment'].filter((r) => firstOf(r)).map((r) => `- ${r}: first due ${fmt(firstOf(r)!.due)}`),
        `- TDS for March: ${rows.find((r) => r.ret === 'TDS payment' && r.period.startsWith('Mar')) ? fmt(rows.find((r) => r.ret === 'TDS payment' && r.period.startsWith('Mar'))!.due) : 'n/a'}`,
        '',
        'Status is a dropdown (Pending / Filed / Not applicable). Rows due within 7 days turn amber, overdue rows red. Every date is editable — edits stay.',
        `**Rule source:** ${ruleCitation('compliance-calendar')}.`,
        '_Due dates are seeded defaults — check notified extensions before relying on them._',
      ].join('\n'),
      explanation: `Create the compliance calendar for FY ${fy}-${String(fy + 1).slice(2)}.`,
      actions,
    };
  },
};

const PRESET_ALIASES: Array<[RegExp, string, string]> = [
  [/\bgst\s*rates?\b/i, 'gst-rates', 'GST rates'],
  [/\btds\s*sections?\b/i, 'tds-sections', 'TDS sections'],
  [/\byes\s*(\/|or)\s*no\b/i, 'yes-no', 'Yes/No'],
];

export const validationPresetLane: DeterministicLane = {
  step: 'P4.2',
  id: 'validation-preset',
  writes: true,
  detect: (m) => /\b(drop-?\s?down|validation|pick\s*list|list)\b/i.test(m) && PRESET_ALIASES.some(([re]) => re.test(m)),
  run: (ctx) => {
    const hit = PRESET_ALIASES.find(([re]) => re.test(ctx.message));
    if (!hit) return null;
    const table = getRuleTable('validation-presets');
    const list = table.data[hit[1]];
    const col = columnsFromMessage(ctx.message)[0];
    if (col === undefined) return { kind: 'answer', answer: `Which column? e.g. *Add a dropdown of ${hit[2]} to column D*.` };
    const first = ctx.headerRowIndex + 2;
    const last = Math.max(lastUsedRow(ctx.values) + 1, first + 199);
    const range = `${columnLetter(col)}${first}:${columnLetter(col)}${last}`;
    return {
      kind: 'actions',
      answer: `Dropdown of **${hit[2]}** on ${range}: ${list.join(', ')}.\n**Rule source:** ${ruleCitation('validation-presets')}.`,
      explanation: `Add a ${hit[2]} dropdown.`,
      actions: [
        {
          type: 'DATA_VALIDATION',
          sheetName: ctx.sheetName,
          range,
          validation: { kind: 'list', listSource: list.map(String), errorTitle: `Invalid ${hit[2]}`, errorMessage: `Pick a value from the ${hit[2]} list.` },
        },
      ] as SheetActionPayload[],
    };
  },
};

export const STAGE4_LANES: DeterministicLane[] = [complianceCalendarLane, validationPresetLane];
