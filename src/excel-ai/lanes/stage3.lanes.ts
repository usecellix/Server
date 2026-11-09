/**
 * Stage 3 deterministic lanes: P3.1 change case, P3.2 fill blanks, P3.3 remove row types,
 * P3.4 lakh/crore display, P3.16 dynamic-array list, P3.18 group-by totals, P3.19 subtotals
 * on a copy, P3.23 trackers + P3.24 template library (incl. P4.7 ageing / compliance).
 */
import { BatchSetOperation, SheetActionPayload } from '../types/sheet-actions.types';
import { GSTIN_FORMAT_REGEX, PAN_REGEX } from '../utils/india-validators.util';
import { DeterministicLane, LaneContext } from './lane.types';
import {
  columnLetter,
  columnsFromMessage,
  findHeader,
  headerNamedInMessage,
  indianNumber,
  isBlank,
  lastUsedColumn,
  lastUsedRow,
  str,
  targetCellFromMessage,
  toNumber,
} from './lane-helpers';

const addr = (col: number, row: number): string => `${columnLetter(col)}${row + 1}`;
const header = (ctx: LaneContext): unknown[] => ctx.values[ctx.headerRowIndex] ?? [];
const dataRowIndexes = (ctx: LaneContext): number[] => {
  const out: number[] = [];
  for (let r = ctx.headerRowIndex + 1; r <= lastUsedRow(ctx.values); r += 1) out.push(r);
  return out;
};
const quoteSheet = (name: string): string => `'${name.replace(/'/g, "''")}'`;


/**
 * WRITE_TABLE writes plain values. Any "=…" cell is moved into a following BATCH_SET so it
 * lands as a real formula (WRITE_TABLE always starts at A1 with the header on row 1).
 */
export function tableWithFormulas(sheetName: string, headers: string[], rows: unknown[][]): SheetActionPayload[] {
  const ops: BatchSetOperation[] = [];
  const plain = rows.map((row, r) =>
    row.map((v, c) => {
      if (typeof v === 'string' && v.startsWith('=')) {
        ops.push({ address: addr(c, r + 1), formula: v });
        return '';
      }
      return v;
    }),
  );
  const out: SheetActionPayload[] = [{ type: 'WRITE_TABLE', sheetName, headers, rows: plain } as SheetActionPayload];
  if (ops.length) out.push({ type: 'BATCH_SET', sheetName, operations: ops } as SheetActionPayload);
  return out;
}

// ---------------------------------------------------------------- P3.1 change case

/** Tokens kept in capitals whatever the mode (CA vocabulary). */
export const PROTECTED_TOKENS = new Set(['LLP', 'HUF', 'GST', 'GSTIN', 'PAN', 'TDS', 'TCS', 'HSN', 'SAC', 'ITC', 'MSME', 'CA', 'NEFT', 'RTGS', 'UPI', 'II', 'III', 'IV']);

export type CaseMode = 'proper' | 'upper' | 'lower' | 'sentence';

export function changeCase(value: string, mode: CaseMode): string {
  if (mode === 'upper') return value.toUpperCase();
  if (mode === 'lower') return value.replace(/\S+/g, (w) => (PROTECTED_TOKENS.has(w.toUpperCase()) ? w.toUpperCase() : w.toLowerCase()));
  if (mode === 'sentence') {
    const lowered = changeCase(value, 'lower');
    return lowered.replace(/^(\s*)(\p{L})/u, (_, s, c) => s + c.toUpperCase());
  }
  return value.replace(/[\p{L}\p{N}][\p{L}\p{N}'.]*/gu, (w) => {
    const bare = w.replace(/\.$/, '').toUpperCase();
    if (PROTECTED_TOKENS.has(bare)) return w.toUpperCase();
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  });
}

const looksLikeIdentifier = (v: unknown): boolean => {
  const s = str(v).trim().toUpperCase();
  return GSTIN_FORMAT_REGEX.test(s) || PAN_REGEX.test(s) || /^[A-Z0-9/-]{6,}$/.test(s) && /\d/.test(s);
};

export const changeCaseLane: DeterministicLane = {
  step: 'P3.1',
  id: 'change-case',
  writes: true,
  detect: (m) =>
    /\b(proper|title|upper|lower|sentence)\s*-?\s*case\b|\b(capitali[sz]e|uppercase|lowercase|all caps)\b/i.test(m) &&
    !/\bsheet name\b/i.test(m),
  run: (ctx) => {
    const mode: CaseMode = /\bupper|all caps/i.test(ctx.message)
      ? 'upper'
      : /\blower/i.test(ctx.message)
        ? 'lower'
        : /\bsentence/i.test(ctx.message)
          ? 'sentence'
          : 'proper';
    const h = header(ctx);
    const explicit = columnsFromMessage(ctx.message);
    const width = lastUsedColumn(ctx.values) + 1;
    const rows = dataRowIndexes(ctx);
    const candidates = explicit.length ? explicit : Array.from({ length: width }, (_, c) => c);
    const skipped: string[] = [];
    const ops: BatchSetOperation[] = [];
    const samples: string[] = [];
    for (const c of candidates) {
      const cells = rows.map((r) => ctx.values[r]?.[c]).filter((v) => typeof v === 'string' && v.trim());
      const idShare = cells.filter(looksLikeIdentifier).length / Math.max(cells.length, 1);
      if (/gstin|\bpan\b|code|ifsc|hsn|sac/i.test(str(h[c])) || idShare >= 0.5) {
        skipped.push(`${columnLetter(c)}${str(h[c]) ? ` (${str(h[c])})` : ''}`);
        continue;
      }
      for (const r of rows) {
        const v = ctx.values[r]?.[c];
        if (typeof v !== 'string' || !v.trim() || (ctx.columnMeta[c]?.formulaRows ?? []).includes(r)) continue;
        const next = changeCase(v, mode);
        if (next === v) continue;
        ops.push({ address: addr(c, r), value: next });
        if (samples.length < 8) samples.push(`${addr(c, r)}: ${v} → ${next}`);
      }
    }
    if (!ops.length) {
      return { kind: 'answer', answer: `Nothing to change — the text is already in ${mode} case${skipped.length ? ` (identifier columns skipped: ${skipped.join(', ')})` : ''}.` };
    }
    return {
      kind: 'actions',
      answer: [
        `Changing **${ops.length}** cells to ${mode} case. Kept in capitals: LLP, HUF, GST, PAN, TDS…`,
        ...(skipped.length ? [`Not touched (identifier columns): ${skipped.join(', ')}.`] : []),
        '',
        ...samples.map((s) => `- ${s}`),
      ].join('\n'),
      explanation: `Change ${ops.length} cells to ${mode} case.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P3.2 fill blanks

export const fillBlanksLane: DeterministicLane = {
  step: 'P3.2',
  id: 'fill-blanks',
  writes: true,
  detect: (m) => /\bfill\b/i.test(m) && /\b(blanks?|empty|empties|gaps?|missing)\b/i.test(m) && !/\bfill\s+down\s+the\s+formula/i.test(m),
  run: (ctx) => {
    const cols = columnsFromMessage(ctx.message);
    const h = header(ctx);
    const col = cols[0] ?? (headerNamedInMessage(h, ctx.message) >= 0 ? headerNamedInMessage(h, ctx.message) : 0);
    const treatSpaces = /\b(treat|count)\s+(the\s+)?spaces?\s+as\s+blanks?\b|\bspaces?\s+as\s+blank/i.test(ctx.message);
    const constant = /\bwith\s+(?:zero|0)\b/i.test(ctx.message)
      ? 0
      : (/\bwith\s+["“']([^"”']+)["”']/i.exec(ctx.message)?.[1] ?? null);
    const method = constant !== null ? 'constant' : /\bbelow\b/i.test(ctx.message) ? 'fromBelow' : 'fromAbove';

    const rows = dataRowIndexes(ctx);
    // Only rows that have something in another column are "inside" the table.
    const inTable = (r: number) => (ctx.values[r] ?? []).some((v, c) => c !== col && !isBlank(v) && str(v).trim() !== '');
    const blank = (v: unknown) => isBlank(v) || (treatSpaces && str(v).trim() === '');
    const spaceOnly: number[] = [];
    const ops: BatchSetOperation[] = [];
    const samples: string[] = [];
    const ordered = method === 'fromBelow' ? [...rows].reverse() : rows;
    let carry: unknown = null;
    for (const r of ordered) {
      const v = ctx.values[r]?.[col];
      if (!inTable(r) && blank(v)) continue;
      if (!blank(v)) {
        if (!treatSpaces && str(v) !== '' && str(v).trim() === '') {
          spaceOnly.push(r);
          continue;
        }
        carry = v;
        continue;
      }
      const fill = method === 'constant' ? constant : carry;
      if (fill === null || fill === undefined) continue;
      ops.push({ address: addr(col, r), value: fill });
      if (samples.length < 8) samples.push(`${addr(col, r)} ← ${str(fill)}`);
    }
    const lines = [
      `Filling **${ops.length}** blank cells in column ${columnLetter(col)} ${method === 'constant' ? `with ${str(constant)}` : method === 'fromBelow' ? 'with the value below' : 'with the value above'}.`,
      ...samples.map((s) => `- ${s}`),
    ];
    if (spaceOnly.length) {
      lines.push('', `Left alone: ${spaceOnly.map((r) => addr(col, r)).join(', ')} contain only spaces (not truly blank). Say *…treat spaces as blank* to fill them too.`);
    }
    if (!ops.length) return { kind: 'answer', answer: `No blank cells to fill in column ${columnLetter(col)}.${spaceOnly.length ? `\n${lines[lines.length - 1]}` : ''}` };
    return {
      kind: 'actions',
      answer: lines.join('\n'),
      explanation: `Fill ${ops.length} blank cells.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P3.3 remove row types

export type RowType = 'data' | 'blank' | 'title' | 'header' | 'header-repeat' | 'subtotal' | 'grand-total';

const TOTAL_RE = /\b(grand\s*total|total)\b/i;
const SUBTOTAL_RE = /\b(sub\s*-?\s*total|subtotal)\b/i;

export function findHeaderRow(values: unknown[][]): number {
  let best = 0;
  let bestScore = -1;
  for (let r = 0; r < Math.min(values.length, 30); r += 1) {
    const row = values[r] ?? [];
    const texts = row.filter((v) => typeof v === 'string' && v.trim() && toNumber(v) === null).length;
    const filled = row.filter((v) => str(v).trim() !== '').length;
    if (filled >= 2 && texts === filled && texts > bestScore) {
      best = r;
      bestScore = texts;
    }
    if (bestScore >= 2 && filled >= 2) break;
  }
  return best;
}

export function classifyRows(values: unknown[][], headerRow: number): RowType[] {
  const head = (values[headerRow] ?? []).map((v) => str(v).trim().toLowerCase());
  const last = lastUsedRow(values);
  return values.slice(0, last + 1).map((row, r) => {
    const cells = (row ?? []).map((v) => str(v).trim());
    const filled = cells.filter(Boolean);
    if (!filled.length) return 'blank';
    if (r === headerRow) return 'header';
    if (r < headerRow) return 'title';
    if (head.length && cells.slice(0, head.length).every((c, i) => c.toLowerCase() === head[i])) return 'header-repeat';
    const text = filled.join(' ');
    if (SUBTOTAL_RE.test(text)) return 'subtotal';
    if (TOTAL_RE.test(text)) return 'grand-total';
    return 'data';
  });
}

export const removeRowTypesLane: DeterministicLane = {
  step: 'P3.3',
  id: 'remove-row-types',
  writes: true,
  detect: (m) =>
    /\b(remove|delete|drop|strip|clean)\b/i.test(m) &&
    /\b(sub\s*-?\s*totals?|subtotal|repeated\s+headers?|header\s+rows?|total\s+rows?|title\s+rows?)\b/i.test(m) &&
    !/\badd\b/i.test(m),
  run: (ctx) => {
    const headerRow = findHeaderRow(ctx.values);
    const types = classifyRows(ctx.values, headerRow);
    const keepTitle = /\bkeep\s+(the\s+)?title\b|\bexcept\s+(the\s+)?title\b/i.test(ctx.message);
    const head = (ctx.values[headerRow] ?? []).map((v) => str(v));
    const width = Math.max(head.length, lastUsedColumn(ctx.values) + 1);
    const count = (t: RowType) => types.filter((x) => x === t).length;
    const dataRows = types.map((t, r) => (t === 'data' ? r : -1)).filter((r) => r >= 0);

    // Tie-out on every numeric column that the grand-total row fills.
    const gtRow = types.lastIndexOf('grand-total');
    const ties: string[] = [];
    let tiesOk = true;
    for (let c = 0; c < width; c += 1) {
      const gt = gtRow >= 0 ? toNumber(ctx.values[gtRow]?.[c]) : null;
      if (gt === null) continue;
      const sum = dataRows.reduce((s, r) => s + (toNumber(ctx.values[r]?.[c]) ?? 0), 0);
      const ok = Math.abs(sum - gt) < 0.005;
      tiesOk = tiesOk && ok;
      ties.push(`${head[c] || columnLetter(c)}: data sum ${indianNumber(sum, sum % 1 ? 2 : 0)} ${ok ? '=' : '≠'} Grand Total ${indianNumber(gt, gt % 1 ? 2 : 0)}`);
    }

    const outRows: unknown[][] = [];
    if (keepTitle) {
      types.forEach((t, r) => {
        if (t === 'title') outRows.push([...(ctx.values[r] ?? [])].slice(0, width));
      });
    }
    const tableRows = dataRows.map((r) => Array.from({ length: width }, (_, c) => ctx.values[r]?.[c] ?? ''));
    const outName = `${ctx.sheetName} - Clean`.slice(0, 31);
    const titleOffset = outRows.length;
    const dateCols = head.map((h, c) => (/date/i.test(h) ? c : -1)).filter((c) => c >= 0);
    const actions: SheetActionPayload[] = [
      { type: 'CREATE_SHEET', sheetName: outName, name: outName, uniqueName: true, relativeTo: ctx.sheetName, position: 'after' },
    ];
    if (keepTitle && titleOffset) {
      actions.push({ type: 'WRITE_TABLE', sheetName: outName, headers: outRows[0].map(str), rows: [...outRows.slice(1), [], head, ...tableRows] });
    } else {
      actions.push({ type: 'WRITE_TABLE', sheetName: outName, headers: head, rows: tableRows });
    }
    const headerAt = keepTitle && titleOffset ? titleOffset + 1 : 0;
    actions.push({ type: 'FORMAT_RANGE', sheetName: outName, range: `A${headerAt + 1}:${columnLetter(width - 1)}${headerAt + 1}`, format: { bold: true } });
    for (const c of dateCols) {
      actions.push({ type: 'FORMAT_RANGE', sheetName: outName, range: `${columnLetter(c)}${headerAt + 2}:${columnLetter(c)}${headerAt + 1 + tableRows.length}`, format: { numberFormat: 'dd-mm-yyyy' } });
    }

    const lines = [
      `Found in **${ctx.sheetName}**: ${count('title')} title, ${count('blank')} blank, ${count('header-repeat')} repeated header, ${count('subtotal')} Sub Total, ${count('grand-total')} Grand Total and **${dataRows.length} data rows**.`,
      `The ${dataRows.length} data rows (with the header${keepTitle ? ' and the title' : ''}) go to a **new sheet "${outName}"** — the original sheet is not changed.`,
      ...(ties.length ? ['', `**Tie-out:** ${ties.join(' · ')}${tiesOk ? '' : ' — ⚠ does not tie, check before relying on it'}`] : ['', 'No Grand Total row found, so there is no tie-out line.']),
      ...(keepTitle ? [] : ['', 'To keep the title row, ask: *…but keep the title*.']),
    ];
    return { kind: 'actions', answer: lines.join('\n'), explanation: `Copy ${dataRows.length} data rows to a clean sheet.`, actions };
  },
};

// ---------------------------------------------------------------- P3.4 lakh / crore display

/** Static per-cell Indian grouping by magnitude — values are never changed. */
export function indianFormatFor(n: number): string {
  const abs = Math.abs(n);
  const dec = Math.abs(n % 1) > 1e-9 ? '.00' : '';
  if (abs >= 1e11) return `##\\,##\\,##\\,##\\,##\\,##0${dec}`;
  if (abs >= 1e9) return `##\\,##\\,##\\,##\\,##0${dec}`;
  if (abs >= 1e7) return `##\\,##\\,##\\,##0${dec}`;
  if (abs >= 1e5) return `##\\,##\\,##0${dec}`;
  return `#,##0${dec}`;
}

export const lakhCroreLane: DeterministicLane = {
  step: 'P3.4',
  id: 'lakh-crore',
  writes: true,
  detect: (m) => /\b(lakhs?|lacs?|crores?|indian\s+(number\s+)?(format|style|grouping|comma))\b/i.test(m) && !/\bchart\b/i.test(m),
  run: (ctx) => {
    const h = header(ctx);
    const cols = columnsFromMessage(ctx.message);
    const col = cols[0] ?? (headerNamedInMessage(h, ctx.message) >= 0 ? headerNamedInMessage(h, ctx.message) : findHeader(h, /amount|value|total|balance|cost/i));
    if (col < 0) return { kind: 'answer', answer: 'Which column? e.g. *Format column A in Indian lakh and crore style*.' };
    const rows = dataRowIndexes(ctx).filter((r) => typeof ctx.values[r]?.[col] === 'number');
    if (!rows.length) return { kind: 'answer', answer: `Column ${columnLetter(col)} has no numbers to format.` };

    const unit = /\bin\s+crores?\b|\bcrores?\s+view\b/i.test(ctx.message) ? 'crore' : /\bin\s+lakhs?\b|\bshow\s+(it\s+)?in\s+lakhs?\b|\blakhs?\s+view\b/i.test(ctx.message) ? 'lakh' : null;
    if (unit) {
      const div = unit === 'lakh' ? 100000 : 10000000;
      const [outCol] = (() => {
        const name = `${str(h[col]) || 'Amount'} (₹ ${unit}s)`;
        const existing = h.findIndex((x) => str(x) === name);
        return [existing >= 0 ? existing : lastUsedColumn(ctx.values) + 1, name];
      })();
      const name = `${str(h[col]) || 'Amount'} (₹ ${unit}s)`;
      const ops: BatchSetOperation[] = [{ address: addr(outCol, ctx.headerRowIndex), value: name }];
      for (const r of rows) ops.push({ address: addr(outCol, r), formula: `=${addr(col, r)}/${div}`, format: { numberFormat: '#,##0.00' } });
      const samples = rows.slice(0, 5).map((r) => ((ctx.values[r]?.[col] as number) / div).toFixed(2));
      return {
        kind: 'actions',
        answer: `Excel number formats can only scale by thousands, so a ₹ ${unit}s view needs a helper column. I'll add **${name}** (column ${columnLetter(outCol)}) as live formulas \`=${addr(col, rows[0])}/${div}\`: ${samples.join(', ')}… Column ${columnLetter(col)} is not changed.`,
        explanation: `Add a ₹ ${unit}s view of column ${columnLetter(col)}.`,
        actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
      };
    }

    const actions: SheetActionPayload[] = rows.map((r) => ({
      type: 'FORMAT_RANGE',
      sheetName: ctx.sheetName,
      range: addr(col, r),
      format: { numberFormat: indianFormatFor(ctx.values[r]?.[col] as number) },
    })) as SheetActionPayload[];
    const samples = rows.slice(0, 6).map((r) => {
      const n = ctx.values[r]?.[col] as number;
      return indianNumber(n, Math.abs(n % 1) > 1e-9 ? 2 : 0);
    });
    return {
      kind: 'actions',
      answer: `Formatting **${rows.length}** numbers in column ${columnLetter(col)} with Indian grouping: ${samples.join(' · ')}. Only the display format changes — the cell values stay exactly the same (check the formula bar). Ask *show in lakhs* for a ₹ lakhs view.`,
      explanation: `Indian lakh/crore format for column ${columnLetter(col)}.`,
      actions,
    };
  },
};

// ---------------------------------------------------------------- P3.16 dynamic-array list

export const dynamicListLane: DeterministicLane = {
  step: 'P3.16',
  id: 'unique-sorted-list',
  writes: true,
  detect: (m) => /\b(unique|distinct)\b/i.test(m) && /\b(list|values|names)\b/i.test(m) && /\b(in|at|into)\s+(cell\s+)?[A-Za-z]{1,2}\d+\b/i.test(m),
  run: (ctx) => {
    const target = targetCellFromMessage(ctx.message);
    if (!target) return null;
    const h = header(ctx);
    const named = headerNamedInMessage(h, ctx.message);
    const col =
      columnsFromMessage(ctx.message)[0] ??
      (named >= 0 ? named : (() => {
        const lower = ctx.message.toLowerCase();
        const idx = h.findIndex((x) => {
          const t = str(x).toLowerCase();
          return t && (lower.includes(t.replace(/s$/, '')) || lower.includes(`${t}s`));
        });
        return idx >= 0 ? idx : 0;
      })());
    const rows = dataRowIndexes(ctx);
    const sorted = /\bsort(ed)?\b/i.test(ctx.message);
    const values = rows.map((r) => ctx.values[r]?.[col]).filter((v) => str(v).trim() !== '');
    const seen = new Map<string, unknown>();
    for (const v of values) if (!seen.has(str(v).toLowerCase())) seen.set(str(v).toLowerCase(), v);
    let list = [...seen.values()];
    if (sorted) list = list.sort((a, b) => str(a).localeCompare(str(b), 'en', { sensitivity: 'base' }));
    const src = `${addr(col, rows[0])}:${addr(col, rows[rows.length - 1])}`;
    const spill = `${addr(target.col, target.row)}:${addr(target.col, target.row + list.length - 1)}`;
    const blockers: string[] = [];
    for (let r = target.row + 1; r < target.row + list.length; r += 1) {
      if (str(ctx.values[r]?.[target.col]).trim() !== '') blockers.push(addr(target.col, r));
    }
    if (blockers.length) {
      return {
        kind: 'answer',
        answer: `⚠ The list needs **${spill}** (${list.length} cells) but ${blockers.join(', ')} ${blockers.length === 1 ? 'is' : 'are'} not empty — the formula would show #SPILL!. Clear ${blockers.length === 1 ? 'it' : 'them'} or pick another cell (e.g. *…in ${columnLetter(lastUsedColumn(ctx.values) + 2)}2*). Nothing was written.`,
      };
    }
    const formula = sorted ? `=SORT(UNIQUE(${src}))` : `=UNIQUE(${src})`;
    if (ctx.dynamicArrays === false) {
      const ops: BatchSetOperation[] = list.map((v, i) => ({ address: addr(target.col, target.row + i), value: v }));
      return {
        kind: 'actions',
        answer: `Your Excel does not support dynamic arrays (UNIQUE/SORT), so I'll write the ${list.length} values as a **static list** in ${spill}: ${list.map(str).join(', ')}. It will not update by itself.`,
        explanation: `Write a static ${sorted ? 'sorted ' : ''}unique list.`,
        actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: ops }] as SheetActionPayload[],
      };
    }
    return {
      kind: 'actions',
      answer: `I'll put \`${formula}\` in ${addr(target.col, target.row)}. It spills into **${spill}** (${list.length} values, all empty now): ${list.map(str).join(', ')}.`,
      explanation: `Unique ${sorted ? 'sorted ' : ''}list in ${addr(target.col, target.row)}.`,
      actions: [{ type: 'BATCH_SET', sheetName: ctx.sheetName, operations: [{ address: addr(target.col, target.row), formula }] }] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P3.18 group-by totals

function parseGroupBy(message: string, h: unknown[]): { value: number; groups: number[] } | null {
  const m = /\b(?:total|sum|subtotal)\s+(?:of\s+)?(?:the\s+)?(.+?)\s+(?:by|per|for each)\s+(.+?)(?:[.;?]|\s+as\s+|$)/i.exec(message);
  if (!m) return null;
  const find = (text: string) => {
    const t = text.trim().toLowerCase();
    return h.findIndex((x) => str(x).trim().toLowerCase() === t);
  };
  const value = find(m[1]);
  const groups = m[2].split(/\s*(?:,|\band\b|&)\s*/i).map(find);
  if (value < 0 || !groups.length || groups.some((g) => g < 0)) return null;
  return { value, groups };
}

export const groupByLane: DeterministicLane = {
  step: 'P3.18',
  id: 'group-by',
  writes: true,
  detect: (m) => /\b(total|sum)\b.+\b(by|per)\b/i.test(m) && !/\bsubtotals?\b/i.test(m) && !/\bchart\b/i.test(m),
  run: (ctx) => {
    const h = header(ctx);
    const spec = parseGroupBy(ctx.message, h);
    if (!spec) return null;
    const asFormulas = /\bsumifs?\b|\blive\b|\bformulas?\b/i.test(ctx.message);
    const rows = dataRowIndexes(ctx);
    const groups = new Map<string, { keys: unknown[]; total: number }>();
    let source = 0;
    for (const r of rows) {
      const n = toNumber(ctx.values[r]?.[spec.value]);
      if (n === null) continue;
      source += n;
      const keys = spec.groups.map((g) => ctx.values[r]?.[g] ?? '');
      const k = keys.map((x) => str(x).toLowerCase()).join('\u0001');
      const cur = groups.get(k) ?? { keys, total: 0 };
      cur.total += n;
      groups.set(k, cur);
    }
    const valueName = str(h[spec.value]);
    const groupNames = spec.groups.map((g) => str(h[g]));
    const outName = `${valueName} by ${groupNames.join(' & ')}`.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31);
    const list = [...groups.values()];
    const first = ctx.headerRowIndex + 2;
    const last = rows.length ? rows[rows.length - 1] + 1 : first;
    const sumifs = (i: number) => {
      const rowNo = i + 2;
      const crit = spec.groups
        .map((g, gi) => `${quoteSheet(ctx.sheetName)}!$${columnLetter(g)}$${first}:$${columnLetter(g)}$${last},${columnLetter(gi)}${rowNo}`)
        .join(',');
      return `=SUMIFS(${quoteSheet(ctx.sheetName)}!$${columnLetter(spec.value)}$${first}:$${columnLetter(spec.value)}$${last},${crit})`;
    };
    const tableRows = list.map((g, i) => [...g.keys, asFormulas ? sumifs(i) : g.total]);
    const totalCol = columnLetter(spec.groups.length);
    tableRows.push([...spec.groups.map((_, i) => (i === 0 ? 'Grand Total' : '')), `=SUM(${totalCol}2:${totalCol}${list.length + 1})`]);
    const grand = list.reduce((s, g) => s + g.total, 0);
    return {
      kind: 'actions',
      answer: [
        `**${valueName} by ${groupNames.join(' and ')}** — ${list.length} groups on a new sheet "${outName}"${asFormulas ? ', as live SUMIFS formulas pointing at the source' : ''}.`,
        ...list.slice(0, 10).map((g) => `- ${g.keys.map(str).join(' / ')}: ${indianNumber(g.total, g.total % 1 ? 2 : 0)}`),
        '',
        `**Tie-out:** groups total ${indianNumber(grand)} = source total ${indianNumber(source)}${Math.abs(grand - source) < 0.005 ? ' ✓' : ' ✗'}.`,
        asFormulas ? '' : 'Ask *…as SUMIFS formulas* for live formulas instead of values.',
      ].filter(Boolean).join('\n'),
      explanation: `Total ${valueName} by ${groupNames.join(', ')}.`,
      actions: [
        { type: 'CREATE_SHEET', sheetName: outName, name: outName, uniqueName: true, relativeTo: ctx.sheetName, position: 'after' },
        ...tableWithFormulas(outName, [...groupNames, `Total ${valueName}`], tableRows),
        { type: 'FORMAT_RANGE', sheetName: outName, range: `A1:${totalCol}1`, format: { bold: true } },
        { type: 'FORMAT_RANGE', sheetName: outName, range: `A${list.length + 2}:${totalCol}${list.length + 2}`, format: { bold: true } },
      ] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P3.19 subtotals on a copy

export const subtotalsLane: DeterministicLane = {
  step: 'P3.19',
  id: 'subtotals',
  writes: true,
  detect: (m) => /\bsub\s*-?\s*totals?\b/i.test(m) && /\b(add|insert|create|put|show)\b/i.test(m),
  run: (ctx) => {
    const h = header(ctx);
    const named = /\b(?:per|by|for each)\s+(?:the\s+)?([A-Za-z][\w ]*?)(?:\s+on\b|\s+in\b|[.,]|$)/i.exec(ctx.message)?.[1];
    const groupCol = named ? h.findIndex((x) => str(x).trim().toLowerCase() === named.trim().toLowerCase()) : columnsFromMessage(ctx.message)[0] ?? 0;
    if (groupCol < 0) return { kind: 'answer', answer: `I couldn't find a column named "${named}".` };
    const numCols = h.map((_, c) => c).filter((c) => c !== groupCol && dataRowIndexes(ctx).some((r) => typeof ctx.values[r]?.[c] === 'number'));
    const width = h.length;
    const rows = dataRowIndexes(ctx).filter((r) => (ctx.values[r] ?? []).some((v) => str(v).trim() !== ''));
    const out: unknown[][] = [];
    const subtotalRows: number[] = [];
    const summary: string[] = [];
    let current: string | null = null;
    let start = 2; // sheet row number of the first data row of the current group (header is row 1)
    let groupTotal = 0;
    const closeGroup = () => {
      if (current === null) return;
      const rowNo = out.length + 2;
      const line: unknown[] = Array.from({ length: width }, () => '');
      line[groupCol] = `${current} Subtotal`;
      for (const c of numCols) line[c] = `=SUBTOTAL(9,${columnLetter(c)}${start}:${columnLetter(c)}${rowNo - 1})`;
      out.push(line);
      subtotalRows.push(rowNo);
      summary.push(`${current}: ${indianNumber(groupTotal)}`);
      groupTotal = 0;
      start = rowNo + 1;
    };
    let grand = 0;
    for (const r of rows) {
      const raw = str(ctx.values[r]?.[groupCol]).trim();
      const g: string = raw || current || '(blank)';
      if (g !== current) {
        closeGroup();
        current = g;
        start = out.length + 2;
      }
      const copy = Array.from({ length: width }, (_, c) => ctx.values[r]?.[c] ?? '');
      out.push(copy);
      const n = numCols.length ? toNumber(ctx.values[r]?.[numCols[0]]) ?? 0 : 0;
      groupTotal += n;
      grand += n;
    }
    closeGroup();
    const lastRow = out.length + 1;
    const grandLine: unknown[] = Array.from({ length: width }, () => '');
    grandLine[groupCol] = 'Grand Total';
    for (const c of numCols) grandLine[c] = `=SUBTOTAL(9,${columnLetter(c)}2:${columnLetter(c)}${lastRow})`;
    out.push(grandLine);
    const outName = `${ctx.sheetName} - Subtotals`.slice(0, 31);
    const boldRanges = [...subtotalRows, out.length + 1].map((rowNo) => ({
      type: 'FORMAT_RANGE',
      sheetName: outName,
      range: `A${rowNo}:${columnLetter(width - 1)}${rowNo}`,
      format: { bold: true },
    }));
    return {
      kind: 'actions',
      answer: [
        `Subtotals per **${str(h[groupCol])}** on a **copy** ("${outName}") — the original sheet is not touched. Blank group cells continue the group above (Tally style).`,
        ...summary.map((s) => `- ${s}`),
        `- **Grand total: ${indianNumber(grand)}** (SUBTOTAL(9,…) ignores the group subtotal rows, so it is not double counted)`,
      ].join('\n'),
      explanation: `Add subtotals per ${str(h[groupCol])} on a copy.`,
      actions: [
        { type: 'CREATE_SHEET', sheetName: outName, name: outName, uniqueName: true, relativeTo: ctx.sheetName, position: 'after' },
        ...tableWithFormulas(outName, h.map(str), out),
        { type: 'FORMAT_RANGE', sheetName: outName, range: `A1:${columnLetter(width - 1)}1`, format: { bold: true } },
        ...boldRanges,
      ] as SheetActionPayload[],
    };
  },
};

// ---------------------------------------------------------------- P3.23 / P3.24 / P4.7 templates

export interface TemplateSpec {
  id: string;
  title: string;
  sheetName: string;
  headers: string[];
  statusColumn?: string;
  statusList?: string[];
  dateColumns?: string[];
  amountColumns?: string[];
  aliases: RegExp;
}

export const TEMPLATES: TemplateSpec[] = [
  {
    id: 'pbc',
    title: 'PBC / document request tracker',
    sheetName: 'PBC Tracker',
    headers: ['Item', 'Requested on', 'Due', 'Received', 'Status', 'Remarks'],
    statusColumn: 'Status',
    statusList: ['Pending', 'Received', 'Partially received', 'Not applicable'],
    dateColumns: ['Requested on', 'Due', 'Received'],
    aliases: /\b(pbc|document request|documents? tracker|request tracker)\b/i,
  },
  {
    id: 'notice',
    title: 'Notice tracker',
    sheetName: 'Notice Tracker',
    headers: ['Authority', 'Notice type', 'DIN', 'Date', 'Reply due', 'Status', 'Remarks'],
    statusColumn: 'Status',
    statusList: ['Open', 'Reply filed', 'Hearing', 'Closed'],
    dateColumns: ['Date', 'Reply due'],
    aliases: /\bnotices?\b/i,
  },
  {
    id: 'gst-recon',
    title: 'GST recon working',
    sheetName: 'GST Recon Working',
    headers: ['Supplier GSTIN', 'Supplier Name', 'Invoice No', 'Invoice Date', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'As per 2B', 'Difference', 'Status', 'Remarks'],
    statusColumn: 'Status',
    statusList: ['Matched', 'Mismatch', 'Missing in 2B', 'Missing in books'],
    dateColumns: ['Invoice Date'],
    amountColumns: ['Taxable Value', 'IGST', 'CGST', 'SGST', 'As per 2B', 'Difference'],
    aliases: /\bgst\s*recon(ciliation)?\s*(working|template)?\b/i,
  },
  {
    id: 'tds',
    title: 'TDS working',
    sheetName: 'TDS Working',
    headers: ['Deductee', 'PAN', 'Section', 'Payment Date', 'Amount Paid', 'TDS Rate', 'TDS Deducted', 'Challan No', 'Deposit Date', 'Remarks'],
    dateColumns: ['Payment Date', 'Deposit Date'],
    amountColumns: ['Amount Paid', 'TDS Deducted'],
    aliases: /\btds\s*(working|template|register)\b/i,
  },
  {
    id: 'far',
    title: 'Fixed asset register',
    sheetName: 'Fixed Asset Register',
    headers: ['Asset', 'Block/Class', 'Date of Purchase', 'Put-to-use Date', 'Cost', 'Additions', 'Deletions', 'Location', 'Tag No.'],
    dateColumns: ['Date of Purchase', 'Put-to-use Date'],
    amountColumns: ['Cost', 'Additions', 'Deletions'],
    aliases: /\b(fixed asset|far)\s*(register|template)\b/i,
  },
  {
    id: 'ageing',
    title: 'Ageing template (P4.7)',
    sheetName: 'Ageing',
    headers: ['Party', 'Invoice No', 'Invoice Date', 'Amount', 'Days Outstanding', '0-30', '31-60', '61-90', '91-180', '>180'],
    dateColumns: ['Invoice Date'],
    amountColumns: ['Amount', '0-30', '31-60', '61-90', '91-180', '>180'],
    aliases: /\bageing|aging\b/i,
  },
  {
    id: 'compliance-tracker',
    title: 'Compliance tracker (P4.7)',
    sheetName: 'Compliance Tracker',
    headers: ['Client', 'Return', 'Period', 'Due date', 'Status', 'Filed on', 'ARN', 'Remarks'],
    statusColumn: 'Status',
    statusList: ['Pending', 'Filed', 'Not applicable'],
    dateColumns: ['Due date', 'Filed on'],
    aliases: /\bcompliance\s+tracker\b/i,
  },
];

export function templateActions(t: TemplateSpec, rows: unknown[][] = []): SheetActionPayload[] {
  const name = t.sheetName;
  const last = columnLetter(t.headers.length - 1);
  const col = (h: string) => columnLetter(t.headers.indexOf(h));
  const actions: SheetActionPayload[] = [
    { type: 'CREATE_SHEET', sheetName: name, name, uniqueName: true },
    { type: 'WRITE_TABLE', sheetName: name, headers: t.headers, rows },
    { type: 'FORMAT_RANGE', sheetName: name, range: `A1:${last}1`, format: { bold: true, fillColor: '#203764', fontColor: '#FFFFFF' } },
    { type: 'FREEZE_PANES', sheetName: name, freezeRows: 1 },
  ] as SheetActionPayload[];
  for (const d of t.dateColumns ?? []) {
    actions.push({ type: 'FORMAT_RANGE', sheetName: name, range: `${col(d)}2:${col(d)}1000`, format: { numberFormat: 'dd-mm-yyyy' } } as SheetActionPayload);
  }
  for (const a of t.amountColumns ?? []) {
    actions.push({ type: 'FORMAT_RANGE', sheetName: name, range: `${col(a)}2:${col(a)}1000`, format: { numberFormat: '#,##,##0.00' } } as SheetActionPayload);
  }
  if (t.statusColumn && t.statusList) {
    actions.push({
      type: 'DATA_VALIDATION',
      sheetName: name,
      range: `${col(t.statusColumn)}2:${col(t.statusColumn)}1000`,
      validation: { kind: 'list', listSource: t.statusList, errorTitle: 'Invalid status', errorMessage: 'Pick a value from the list.' },
    } as SheetActionPayload);
  }
  actions.push({ type: 'AUTOFIT_COLUMNS', sheetName: name } as SheetActionPayload);
  return actions;
}

export const templatesLane: DeterministicLane = {
  step: 'P3.23/P3.24',
  id: 'templates',
  writes: true,
  detect: (m) =>
    /\b(show|list|what)\b.*\btemplates?\b/i.test(m) ||
    (/\b(create|make|build|new|set up|setup|generate|use)\b/i.test(m) &&
      /\b(tracker|template|working|register)\b/i.test(m) &&
      TEMPLATES.some((t) => t.aliases.test(m))),
  run: (ctx) => {
    if (/\b(show|list|what)\b.*\btemplates?\b/i.test(ctx.message) && !/\b(create|make|use)\b/i.test(ctx.message)) {
      return {
        kind: 'answer',
        answer: [
          '**Starter templates** — ask *Create a <name>* and I will preview it first:',
          ...TEMPLATES.map((t) => `- **${t.title}** — ${t.headers.join(' · ')}`),
          '',
          'Your own: select a sheet and ask *Save this sheet as a template* (only the headers are saved, never data).',
        ].join('\n'),
      };
    }
    if (/\bcompliance\s+calendar\b/i.test(ctx.message)) return null;
    const t = TEMPLATES.find((x) => x.aliases.test(ctx.message));
    if (!t) return null;
    return {
      kind: 'actions',
      answer: `I'll create **${t.sheetName}** — ${t.headers.join(' · ')}.${t.statusList ? ` **${t.statusColumn}** gets a dropdown: ${t.statusList.join(' / ')}.` : ''} Dates dd-mm-yyyy. If the name is taken a "(2)" copy is made.`,
      explanation: `Create ${t.title}.`,
      actions: templateActions(t),
    };
  },
};

export const STAGE3_LANES: DeterministicLane[] = [
  templatesLane,
  removeRowTypesLane,
  subtotalsLane,
  changeCaseLane,
  fillBlanksLane,
  lakhCroreLane,
  dynamicListLane,
  groupByLane,
];
