import {
  SMOKE_CASES,
  SmokeCase,
  sheetsUnderColumnCheck,
  validateCases,
} from './smoke-cases';

/**
 * End-to-end smoke test for the 12-month booking-ledger prompt — the live
 * failure behind TASKS.md #259-#263.
 *
 * Drives the REAL backend the way the task pane does: one POST to
 * /excel-ai/conversation, then a POST to /conversation/continue per
 * `wave_ready`, accepting every wave, until the run ends. Every action the run
 * emits is replayed against a virtual workbook so the two things that actually
 * broke in Excel can be asserted without Excel:
 *
 *   1. a write landing on a sheet no action ever created ("The requested
 *      resource doesn't exist" at Accept — what made Accept look dead), and
 *   2. a formula referencing a sheet that never gets created (#REF!).
 *
 *   CELLIX_EVAL_BASE_URL=http://localhost:4011 npx ts-node eval/ledger-smoke.ts
 */

const BASE_URL = process.env.CELLIX_EVAL_BASE_URL ?? 'http://localhost:4011';
const TOKEN = process.env.CELLIX_EVAL_BYPASS_TOKEN ?? 'usecase-probe';
const MAX_WAVES = Number(process.env.MAX_WAVES ?? 25);

// Which of the eight prompts to run: a case id, or 'all'. Defaults to the
// booking ledger, so every existing invocation behaves exactly as before.
const CASE_SELECTOR = process.env.CELLIX_SMOKE_CASE ?? 'ledger';
const DRY_RUN = process.env.CELLIX_SMOKE_DRY_RUN === '1';

const PROMPT =
  'i like to have multiple sheets for all months in a year, and need a main sheet it has all the ' +
  'details of the remaining sheets, in the main sheet i need to have dashboard also, my need to ' +
  'record payments and related things ,which all month sheets include Unit No, Guest, Guest name, ' +
  'check in, check out, Rate per night, total amount, source, payment status, bank account';

/**
 * An empty workbook — exactly what the user ran against.
 *
 * Shape matches `usecase-fixture.ts` exactly: the add-in sends `sheetName`,
 * not `name`. Getting that wrong is what surfaced the crash fixed in #263.
 */
function emptyWorkbook() {
  return {
    sheetData: [[]],
    workbookContext: {
      activeSheet: 'Sheet1',
      sheets: [
        {
          sheetName: 'Sheet1',
          usedRange: 'A1',
          rowCount: 0,
          colCount: 0,
          headers: [] as string[],
          sampleData: [] as unknown[][],
          isHidden: false,
        },
      ],
    },
  };
}

interface SseEvent {
  event: string;
  data: Record<string, unknown>;
}

function parseSse(raw: string): SseEvent[] {
  const out: SseEvent[] = [];
  for (const frame of raw.split('\n\n')) {
    const event = frame
      .split('\n')
      .find((l) => l.startsWith('event:'))
      ?.slice(6)
      .trim();
    const data = frame
      .split('\n')
      .filter((l) => l.startsWith('data:'))
      .map((l) => l.slice(5).trim())
      .join('');
    if (!event || !data) continue;
    try {
      out.push({ event, data: JSON.parse(data) as Record<string, unknown> });
    } catch {
      /* non-JSON frame — ignore */
    }
  }
  return out;
}

type Action = Record<string, unknown> & { type: string };

/** An indented list under a report line, or nothing when there is none. */
const bullets = (xs: string[]): string => (xs.length ? '\n   ' + xs.join('\n   ') : '');

/** Sheet names an action creates. */
function createdSheet(a: Action): string | null {
  if (a.type !== 'ADD_SHEET' && a.type !== 'CREATE_SHEET') return null;
  const name = String(a.name ?? a.sheetName ?? '').trim();
  return name || null;
}

/** Sheet an action writes to, when it needs that sheet to already exist. */
function writesToSheet(a: Action): string | null {
  if (createdSheet(a)) return null;
  const structural = new Set(['DELETE_SHEET', 'RENAME_SHEET', 'COPY_SHEET', 'MOVE_SHEET']);
  if (structural.has(a.type)) return null;
  const name = String(a.sheetName ?? a.destSheet ?? '').trim();
  return name || null;
}

const SHEET_REF = /(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!\$?[A-Za-z]{1,3}\$?\d*/g;

/** Sheets a formula in this action reads from. */
function referencedSheets(a: Action): string[] {
  const texts: string[] = [];
  const collect = (v: unknown) => {
    if (typeof v === 'string' && v.startsWith('=')) texts.push(v);
  };
  collect(a.formula);
  if (Array.isArray(a.operations)) {
    for (const op of a.operations as Array<Record<string, unknown>>) {
      collect(op.formula);
      collect(op.value);
    }
  }
  const names = new Set<string>();
  for (const text of texts) {
    for (const m of text.matchAll(SHEET_REF)) {
      const n = (m[1]?.replace(/''/g, "'") ?? m[2] ?? '').trim();
      if (n) names.add(n);
    }
  }
  return [...names];
}

/**
 * Deliberately `node:http` rather than `fetch`. Node's fetch (undici) applies
 * a 5-minute body timeout that cannot be configured without the `undici`
 * package, and a legitimate wave here can run to the server's own 480s budget:
 * a real 12-month run died with `UND_ERR_BODY_TIMEOUT` mid-wave, so the
 * harness was unable to measure the very case it exists to measure.
 */
async function post(path: string, body: unknown): Promise<SseEvent[]> {
  const http = await import('node:http');
  const payload = JSON.stringify(body);
  const url = new URL(`${BASE_URL}${path}`);

  return new Promise<SseEvent[]>((resolve, reject) => {
    const req = http.request(
      {
        hostname: url.hostname,
        port: url.port,
        path: url.pathname + url.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(payload),
          'x-cellix-eval-bypass': TOKEN,
        },
      },
      (res) => {
        let text = '';
        res.setEncoding('utf8');
        res.on('data', (chunk: string) => {
          text += chunk;
        });
        res.on('end', () => {
          if ((res.statusCode ?? 0) < 200 || (res.statusCode ?? 0) >= 300) {
            reject(new Error(`${path} -> HTTP ${res.statusCode}: ${text.slice(0, 400)}`));
            return;
          }
          resolve(parseSse(text));
        });
      },
    );
    // No socket/response timeout: the server owns the budget, not the client.
    req.setTimeout(0);
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

async function runCase(testCase: SmokeCase): Promise<boolean> {
  const started = Date.now();
  const fixture = emptyWorkbook();

  // Unique per run — omitting workbookId let repeated invocations resolve to
  // the same underlying conversation record, which could carry stale history.
  const workbookId = `wb_smoke_${testCase.id}_${Date.now()}`;
  console.log('');
  console.log(`════════ ${testCase.label} [${testCase.id}] ════════`);
  console.log(`[smoke] POST /excel-ai/conversation  (${BASE_URL})`);
  let events = await post('/excel-ai/conversation', {
    message: testCase.prompt,
    sheetData: fixture.sheetData,
    workbookContext: fixture.workbookContext,
    mode: 'action',
    workbookId,
  });

  const allActions: Action[] = [];
  const waveSizes: number[] = [];
  let runId: string | undefined;
  let clarification: string | undefined;
  let waves = 0;
  /**
   * An `error` frame, or a run that stops before its last wave, must FAIL.
   * Run 4 ended at wave 7 of 8 on an out-of-credits error and still printed
   * PASS, because every check it had passed on the part that got built —
   * the harness committing the exact silent-failure the plan’s §6 item 3
   * calls the one that matters. A partial build is not a pass.
   */
  const errors: string[] = [];

  const absorb = (evts: SseEvent[]) => {
    for (const { event, data } of evts) {
      if (event === 'actions' && Array.isArray(data.actions)) {
        const batch = data.actions as Action[];
        allActions.push(...batch);
        waveSizes.push(batch.length);
      }
      if (event === 'clarification' && typeof data.question === 'string') {
        clarification = data.question;
      }
      if (event === 'wave_ready' && typeof data.runId === 'string') {
        runId = data.runId;
      }
      if (event === 'error') {
        const message = String(data.message ?? JSON.stringify(data)).slice(0, 300);
        errors.push(message);
        console.log(`[smoke] ERROR event: ${message}`);
      }
    }
  };

  absorb(events);

  while (
    runId &&
    events.some((e) => e.event === 'wave_ready' && e.data.hasMore !== false) &&
    waves < MAX_WAVES
  ) {
    waves += 1;
    const readyMeta = events.find((e) => e.event === 'wave_ready')?.data;
    process.stdout.write(
      `[smoke] wave ${waves} → accept … (waveIndex=${readyMeta?.waveIndex}, waveTotal=${readyMeta?.waveTotal}, hasMore=${readyMeta?.hasMore})\n`,
    );
    events = await post('/excel-ai/conversation/continue', { runId, decision: 'accepted' });
    absorb(events);
  }
  const finalMeta = events.find((e) => e.event === 'wave_ready')?.data;
  if (finalMeta) {
    process.stdout.write(
      `[smoke] loop ended (waveIndex=${finalMeta.waveIndex}, waveTotal=${finalMeta.waveTotal}, hasMore=${finalMeta.hasMore})\n`,
    );
  }

  // ---- Replay every action against a virtual workbook, in emission order ----
  const existing = new Set<string>(['sheet1']);
  const key = (n: string) => n.trim().toLowerCase();
  const createOrder: string[] = [];
  const writesBeforeCreate: Array<{ type: string; sheet: string; wave: number }> = [];

  // Which wave each action was emitted in. A write landing before its create
  // is only actionable if you know WHICH step did it — without this the
  // report names the action types and leaves the attribution to guesswork.
  const waveOfAction: number[] = [];
  waveSizes.forEach((size, wave) => {
    for (let i = 0; i < size; i += 1) waveOfAction.push(wave + 1);
  });
  const refsToMissing: Array<{ sheet: string; missing: string }> = [];

  for (const [index, action] of allActions.entries()) {
    const created = createdSheet(action);
    if (created) {
      if (!existing.has(key(created))) createOrder.push(created);
      existing.add(key(created));
      continue;
    }
    const target = writesToSheet(action);
    if (target && !existing.has(key(target))) {
      writesBeforeCreate.push({ type: action.type, sheet: target, wave: waveOfAction[index] ?? -1 });
    }
  }

  // Formula refs are judged against the FINAL sheet set: a formula may legally
  // reference a sheet a later wave creates.
  for (const action of allActions) {
    for (const ref of referencedSheets(action)) {
      if (!existing.has(key(ref))) {
        refsToMissing.push({ sheet: String(action.sheetName ?? '?'), missing: ref });
      }
    }
  }

  const required = testCase.requiredSheets;
  const missingRequired = required.filter((name) => !existing.has(key(name)));
  const summarySheet = testCase.summarySheet;
  const hasSummary = !summarySheet || existing.has(key(summarySheet));

  const uniq = <T>(xs: T[]) => [...new Set(xs.map((x) => JSON.stringify(x)))].map((s) => JSON.parse(s) as T);

  console.log('\n──────── RESULT ────────');
  console.log(`elapsed        : ${((Date.now() - started) / 1000).toFixed(1)}s`);
  console.log(`waves accepted : ${waves}  (action counts: ${waveSizes.join(', ')})`);
  console.log(`total actions  : ${allActions.length}`);
  if (clarification) console.log(`clarification  : ${clarification}`);
  console.log(`sheets created : ${createOrder.length} → ${createOrder.join(', ') || '(none)'}`);
  if (summarySheet) console.log(`${summarySheet} created : ${hasSummary ? 'YES' : 'NO  ← #262'}`);
  console.log(`required sheets missing : ${missingRequired.length ? missingRequired.join(', ') : 'none'}`);

  const uniqueWrites = uniq(writesBeforeCreate);
  console.log(
    `writes onto a sheet that does not exist yet : ${uniqueWrites.length}` +
      (uniqueWrites.length
        ? `\n   ${uniqueWrites.map((w) => `wave ${w.wave}: ${w.type} → ${w.sheet}`).join('\n   ')}`
        : ''),
  );

  const uniqueRefs = uniq(refsToMissing);
  console.log(
    `formulas referencing a never-created sheet  : ${uniqueRefs.length}` +
      (uniqueRefs.length
        ? `\n   ${uniqueRefs.map((r) => `${r.sheet} → ${r.missing}`).join('\n   ')}`
        : ''),
  );

  // TASKS.md #265 — a column width below the readability floor means real
  // text (header names) render as 1-2 clipped characters, reading as a
  // broken/blank sheet even though every value is correct.
  const MIN_READABLE_WIDTH = 40;
  const tooNarrow = allActions.filter(
    (a) => a.type === 'SET_COLUMN_WIDTH' && typeof a.width === 'number' && (a.width as number) < MIN_READABLE_WIDTH,
  );
  console.log(
    `column widths below ${MIN_READABLE_WIDTH}pt (unreadable)         : ${tooNarrow.length}` +
      (tooNarrow.length
        ? `\n   ${tooNarrow.map((a) => `${a.sheetName} col ${a.col}: ${a.width}pt`).join('\n   ')}`
        : ''),
  );

  // ---- §3's two missing checks: header text, and derived-column formulas ----
  const headerRows = replayHeaderRows(allActions);
  const withFormulas = formulaColumns(allActions);
  const headerFailures: string[] = [];
  const derivedFailures: string[] = [];

  for (const month of sheetsUnderColumnCheck(testCase)) {
    // A month that was never created is already reported above; reporting it
    // twice would read as two separate faults.
    if (!existing.has(key(month))) continue;

    const headers = headerList(headerRows.get(key(month)));
    if (headers.length === 0) {
      headerFailures.push(`${month}: no header row written`);
      continue;
    }
    const defects = headerDefects(headers, testCase.columns);
    if (defects.length) {
      headerFailures.push(`${month}: ${defects.join('; ')}  — got [${headers.join(' | ')}]`);
    }

    for (const derivedColumn of testCase.derivedColumns) {
      const derivedAt = headers.findIndex((h) => norm(h) === norm(derivedColumn));
      if (derivedAt === -1) continue; // already counted as a header defect
      if (!withFormulas.get(key(month))?.has(derivedAt + 1)) {
        derivedFailures.push(
          `${month}: "${derivedColumn}" (column ${derivedAt + 1}) carries no formula`,
        );
      }
    }
  }

  console.log(
    `header rows not matching the prompt's columns : ${headerFailures.length}` +
      bullets(headerFailures),
  );
  console.log(
    `derived columns carrying no formula           : ${derivedFailures.length}` +
      bullets(derivedFailures),
  );

  const lastMeta = events.find((e) => e.event === 'wave_ready')?.data;
  const waveTotal = Number(lastMeta?.waveTotal ?? 0);
  const finishedAllWaves = errors.length === 0 && (waveTotal === 0 || waves >= waveTotal);
  console.log(
    `run errors                                   : ${errors.length}` + bullets(errors),
  );
  console.log(
    `all planned waves ran                        : ${finishedAllWaves ? 'YES' : `NO (${waves} of ${waveTotal})`}`,
  );

  const ok =
    hasSummary &&
    missingRequired.length === 0 &&
    uniqueWrites.length === 0 &&
    uniqueRefs.length === 0 &&
    tooNarrow.length === 0 &&
    headerFailures.length === 0 &&
    derivedFailures.length === 0 &&
    finishedAllWaves;
  console.log(`\nVERDICT: ${ok ? 'PASS' : 'FAIL'}`);
  return ok;
}


// ─────────────────────────────────────────────────────────────────────────────
// The two checks §3 of LONG_PROMPT_RELIABILITY_PLAN.md says this harness needs
// and has never had: header text vs the prompt's OWN column list, and formulas
// actually present in the derived column. Without them the harness passes a
// sheet whose headers are wrong — which is exactly how #283/#284 reached a
// live workbook while this file printed PASS.
// ─────────────────────────────────────────────────────────────────────────────

/** The column list the PROMPT itself spells out, in the order it spells it. */
const PROMPT_COLUMNS = [
  'Unit No', 'Guest', 'Guest name', 'check in', 'check out',
  'Rate per night', 'total amount', 'source', 'payment status', 'bank account',
];

/** The derived column: its cells must carry a formula, not a typed number. */
const DERIVED_COLUMN = 'total amount';

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');

/** `A1` / `$AB$12` -> { col: 1-based, row: 1-based }. */
function parseA1(address: string): { col: number; row: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(address.trim());
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

/**
 * Row 1 of every sheet, replayed from the actions in emission order — the same
 * way Excel would end up with it. Later writes to a cell win, which is what
 * makes a placeholder-then-overwrite sequence read correctly.
 */
function replayHeaderRows(actions: Action[]): Map<string, Map<number, string>> {
  const rows = new Map<string, Map<number, string>>();
  const rowFor = (sheet: string) => {
    const k = sheet.trim().toLowerCase();
    if (!rows.has(k)) rows.set(k, new Map());
    return rows.get(k)!;
  };

  for (const a of actions) {
    const sheet = String(a.sheetName ?? a.name ?? '').trim();
    if (!sheet) continue;

    if (a.type === 'BATCH_SET' && Array.isArray(a.operations)) {
      for (const op of a.operations as Array<Record<string, unknown>>) {
        const at = parseA1(String(op.address ?? ''));
        if (at?.row === 1 && typeof op.value === 'string') rowFor(sheet).set(at.col, op.value);
      }
      continue;
    }
    // SET_CELL carries 0-based row/col (normalizeExecutorOutput has already run).
    if (a.type === 'SET_CELL' && Number(a.row) === 0 && typeof a.value === 'string') {
      rowFor(sheet).set(Number(a.col) + 1, a.value);
      continue;
    }
    // An inserted column contributes a header too; position is resolved by the
    // client, so it is appended past the current span rather than guessed at.
    if (a.type === 'INSERT_COLUMN' && typeof a.columnName === 'string') {
      const r = rowFor(sheet);
      r.set(Math.max(0, ...r.keys()) + 1, a.columnName);
    }
  }
  return rows;
}

/** Header row in column order, trimmed of the trailing empties. */
function headerList(row: Map<number, string> | undefined): string[] {
  if (!row || row.size === 0) return [];
  const last = Math.max(...row.keys());
  const out: string[] = [];
  for (let c = 1; c <= last; c += 1) out.push((row.get(c) ?? '').trim());
  while (out.length && out[out.length - 1] === '') out.pop();
  return out;
}

/**
 * Every prompt column present, in the prompt's own order. Extra computed
 * columns between them are allowed (Phase 1's rule); a MISSING one, or a
 * `Column7` placeholder standing in for one, is the #283/#284 failure.
 */
function headerDefects(headers: string[], promptColumns: string[]): string[] {
  const defects: string[] = [];
  const placeholders = headers.filter((h) => /^column\s*\d+$/i.test(h));
  if (placeholders.length) defects.push(`placeholder headers: ${placeholders.join(', ')}`);
  if (headers.some((h) => h === '')) defects.push('gap (empty cell) inside the header row');

  const seen = headers.map(norm);
  let at = 0;
  const missing: string[] = [];
  for (const want of promptColumns) {
    const found = seen.indexOf(norm(want), at);
    if (found === -1) missing.push(want);
    else at = found + 1;
  }
  if (missing.length) {
    defects.push(
      seen.some((h) => missing.some((m) => h === norm(m)))
        ? `prompt columns out of order: ${missing.join(', ')}`
        : `prompt columns missing: ${missing.join(', ')}`,
    );
  }
  return defects;
}

/** Sheet -> the 1-based columns that receive a formula anywhere below row 1. */
function formulaColumns(actions: Action[]): Map<string, Set<number>> {
  const out = new Map<string, Set<number>>();
  const add = (sheet: string, col: number) => {
    const k = sheet.trim().toLowerCase();
    if (!out.has(k)) out.set(k, new Set());
    out.get(k)!.add(col);
  };

  for (const a of actions) {
    const sheet = String(a.sheetName ?? '').trim();
    if (!sheet) continue;

    if (Array.isArray(a.operations)) {
      for (const op of a.operations as Array<Record<string, unknown>>) {
        const text = op.formula ?? op.value;
        if (typeof text !== 'string' || !text.startsWith('=')) continue;
        const at = parseA1(String(op.address ?? ''));
        if (at && at.row > 1) add(sheet, at.col);
      }
    }
    if (typeof a.formula === 'string' && a.formula.startsWith('=')) {
      const raw = String(a.range ?? a.address ?? '').split(':')[0];
      const at = parseA1(raw);
      if (at && at.row > 1) add(sheet, at.col);
      else if (a.col !== undefined && Number(a.row) > 0) add(sheet, Number(a.col) + 1);
    }
  }
  return out;
}

/**
 * Runs one case, or every case with CELLIX_SMOKE_CASE=all.
 *
 * The suite fails if ANY case fails — LONG_PROMPT_RELIABILITY_PLAN.md §6
 * item 1 asks for 8 of 8, not a majority.
 */
async function main(): Promise<void> {
  const problems = validateCases();
  if (problems.length > 0) {
    console.log('');
    console.log('CASE DEFINITIONS INVALID:' + bullets(problems));
    process.exit(2);
  }
  if (DRY_RUN) {
    console.log('');
    console.log(`${SMOKE_CASES.length} case definitions valid (no server contacted):`);
    for (const testCase of SMOKE_CASES) {
      console.log(
        `   ${testCase.id.padEnd(11)} ${String(testCase.requiredSheets.length).padStart(2)} sheets, ` +
          `${testCase.columns.length} columns, derived: ${testCase.derivedColumns.join(', ')}` +
          `${testCase.summarySheet ? ` + ${testCase.summarySheet}` : ''}`,
      );
    }
    process.exit(0);
  }

  const selected =
    CASE_SELECTOR === 'all'
      ? SMOKE_CASES
      : SMOKE_CASES.filter((c) => c.id === CASE_SELECTOR);
  if (selected.length === 0) {
    console.log(
      `Unknown case "${CASE_SELECTOR}". Known: ${SMOKE_CASES.map((c) => c.id).join(', ')}, or "all".`,
    );
    process.exit(2);
  }

  const results: Array<{ id: string; ok: boolean }> = [];
  for (const testCase of selected) {
    // Deliberately sequential: these runs are heavy, and running them in
    // parallel would reintroduce the very provider contention §1 blamed.
    results.push({ id: testCase.id, ok: await runCase(testCase) });
  }

  if (results.length > 1) {
    const passed = results.filter((r) => r.ok);
    console.log('');
    console.log(`════════ SUITE: ${passed.length} of ${results.length} passed ════════`);
    for (const result of results) {
      console.log(`   ${result.ok ? 'PASS' : 'FAIL'}  ${result.id}`);
    }
  }
  process.exit(results.every((r) => r.ok) ? 0 : 1);
}

main().catch((error) => {
  console.error('[smoke] fatal:', error);
  process.exit(2);
});
