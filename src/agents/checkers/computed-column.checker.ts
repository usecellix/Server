import { Injectable } from '@nestjs/common';
import { Action, isDeterministicStep, VerifierIssue } from '../types/agent.types';
import { buildCheckerResult, CheckerResult, SubtaskActionSlice } from './checker.types';

/**
 * A data-entry template whose computed columns are plain text — TASKS.md #270.
 *
 * The planner's BUILD QUALITY rule (a) has always told plans to add derived
 * columns "as a formula column", and the ledger rules now spell the formulas
 * out. A live run still shipped 12 month sheets where "Total Amount" was a
 * header and nothing else: the user typed a guest, two dates and a rate, and
 * the cell stayed blank. Every existing checker passed it, because each one
 * grades what WAS emitted — none asks whether a column the subtask itself
 * promised to compute actually got a formula.
 *
 * So this is the deterministic half of that rule: if a subtask writes a header
 * whose name is inherently arithmetic (Total Amount, Nights, Balance Due …)
 * and emits no formula anywhere for that sheet, fail it. The retry then goes
 * back to the Executor naming the exact column that needs one.
 *
 * Deliberately narrow:
 * - Only fires when the subtask actually WRITES the header (a formatting or
 *   validation pass over an existing sheet promises no computation).
 * - Only for headers that are arithmetic BY DEFINITION. "Source", "Guest",
 *   "Bank Account" are inputs and must never be flagged; a false positive here
 *   costs a real retry on work that was already correct.
 * - Any formula for that sheet satisfies it. Matching a formula to its column
 *   would need the emitted formula's target column resolved against the header
 *   row, and the failure this exists to catch is the total absence of them.
 */

/**
 * Header labels that are a calculation, not an input. Matched on normalized
 * text so "Total Amount", "total amount" and "TOTAL AMOUNT" all hit.
 *
 * "Amount Received" is deliberately ABSENT: it is money the user records, not
 * a derivation — the same reason "Rate Per Night" is absent.
 */
const COMPUTED_HEADER_PATTERNS: Array<{ label: string; test: RegExp }> = [
  { label: 'Nights', test: /^nights?$/ },
  { label: 'Total Amount', test: /^total\s*(amount|amt|value|price|cost)$/ },
  { label: 'Balance Due', test: /^(balance\s*(due)?|amount\s*due|outstanding)$/ },
  { label: 'Days Overdue', test: /^days?\s*overdue$/ },
  { label: 'Gross Pay', test: /^gross\s*(pay|salary|amount)$/ },
  { label: 'Line Total', test: /^(line|row)\s*total$/ },
  { label: 'Subtotal', test: /^sub\s*total$/ },
];

const normalizeHeader = (value: unknown): string =>
  typeof value === 'string' ? value.trim().toLowerCase().replace(/\s+/g, ' ') : '';

/** Every header-ish string this subtask wrote, with the sheet it landed on. */
function collectWrittenHeaders(actions: Action[]): Array<{ sheetName: string; label: string }> {
  const written: Array<{ sheetName: string; label: string }> = [];

  for (const action of actions) {
    const record = action as unknown as Record<string, unknown>;
    const sheetName = String(record.sheetName ?? '');
    if (!sheetName) continue;

    if (action.type === 'BATCH_SET' && Array.isArray(record.operations)) {
      for (const op of record.operations as Array<Record<string, unknown>>) {
        // A cell carrying a formula is already computed — it is the very thing
        // this checker wants to see, never a bare header to complain about.
        if (op.formula) continue;
        written.push({ sheetName, label: normalizeHeader(op.value) });
      }
      continue;
    }

    if (action.type === 'SET_CELL' && !record.formula) {
      written.push({ sheetName, label: normalizeHeader(record.value) });
    }
  }

  return written.filter((entry) => entry.label.length > 0);
}

/** Sheets this subtask emitted at least one formula for. */
function sheetsWithFormulas(actions: Action[]): Set<string> {
  const sheets = new Set<string>();

  for (const action of actions) {
    const record = action as unknown as Record<string, unknown>;
    const sheetName = String(record.sheetName ?? '');
    if (!sheetName) continue;

    if (typeof record.formula === 'string' && record.formula.trim().startsWith('=')) {
      sheets.add(sheetName.toLowerCase());
      continue;
    }
    if (action.type === 'BATCH_SET' && Array.isArray(record.operations)) {
      const hasFormula = (record.operations as Array<Record<string, unknown>>).some(
        (op) => typeof op.formula === 'string' && op.formula.trim().startsWith('='),
      );
      if (hasFormula) sheets.add(sheetName.toLowerCase());
    }
    // FILL_DOWN propagates an existing formula rather than carrying one, so a
    // subtask using it has satisfied the intent too.
    if (action.type === 'FILL_DOWN') {
      sheets.add(sheetName.toLowerCase());
    }
  }

  return sheets;
}

@Injectable()
export class ComputedColumnChecker {
  check(states: SubtaskActionSlice[]): CheckerResult {
    const subtaskResults = states.map((state) => {
      // LONG_PROMPT_RELIABILITY_PLAN.md — a deterministic header/table step
      // (split out by `splitSpecPinnedSubtasks`) writes the header row on
      // purpose with no formula of its own; the formula is the DEPENDENT
      // subtask's job, one wave later. Grading this one against a rule meant
      // for a single subtask that does both would fail it by construction,
      // forever. TASKS.md #280.
      if (isDeterministicStep(state.subtask)) {
        return {
          subtaskId: state.subtask.id,
          passed: true,
          feedback: 'Computed-column checks passed',
          issues: [],
        };
      }

      // Headers this subtask wrote itself, PLUS the header row a
      // deterministic step built for it one wave earlier. TASKS.md #298:
      // after Phase 1.5 split header-writing away from formula-writing, a
      // "rest" step writes no headers at all, so `collectWrittenHeaders`
      // returned nothing and this checker silently had no opinion on the
      // very subtask whose job the formulas now are. A live run shipped
      // April with validations, formats and widths but no formula at all,
      // reporting `completed: true` — and every net in the pipeline passed
      // it. The pinned row is the same information the header step used, so
      // pairing it with this subtask’s target sheet restores the check
      // the split removed, and does it in time to drive a retry.
      const pinnedRow = state.subtask.resolvedHeaderRow ?? state.subtask.expectedHeaders ?? [];
      const pinnedSheet = state.subtask.targetSheet?.trim() ?? '';
      const headers = [
        ...collectWrittenHeaders(state.actions),
        ...(pinnedSheet
          ? pinnedRow.map((label) => ({ sheetName: pinnedSheet, label: normalizeHeader(label) }))
          : []),
      ].filter((entry) => entry.label.length > 0);
      const formulaSheets = sheetsWithFormulas(state.actions);

      const issues: VerifierIssue[] = [];
      const reported = new Set<string>();

      for (const { sheetName, label } of headers) {
        if (formulaSheets.has(sheetName.toLowerCase())) continue;

        const match = COMPUTED_HEADER_PATTERNS.find((pattern) => pattern.test.test(label));
        if (!match) continue;

        const key = `${sheetName.toLowerCase()}::${match.label}`;
        if (reported.has(key)) continue;
        reported.add(key);

        issues.push({
          severity: 'error',
          subtaskId: state.subtask.id,
          actionIndex: undefined,
          description:
            `Sheet "${sheetName}" got a "${match.label}" column header but this subtask emitted no ` +
            `formula for that sheet — the column can only ever be blank, so the template does not compute.`,
          suggestion:
            `Write "${match.label}"'s formula into the first DATA row (the row under the header) and guard ` +
            `it against blanks, e.g. =IF(OR(D2="",E2=""),"",E2-D2). Inside a Table that becomes the ` +
            `column's formula and fills every row the user types afterwards.`,
        });
      }

      const passed = issues.length === 0;
      return {
        subtaskId: state.subtask.id,
        passed,
        feedback: passed
          ? 'Computed-column checks passed'
          : issues.map((issue) => issue.description).join('; '),
        issues,
      };
    });

    return buildCheckerResult(subtaskResults);
  }
}
