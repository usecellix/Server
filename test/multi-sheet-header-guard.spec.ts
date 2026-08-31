import {
  groupActionsBySheet,
  resolveSheetHeaderStates,
  sheetKeyOf,
  sheetsCreatedInBatch,
} from '../src/excel-ai/utils/sheet-header-state.util';
import { SheetActionPayload } from '../src/excel-ai/types/sheet-actions.types';
import { WorkbookContext } from '../src/types/cellix.types';

/**
 * TASKS.md #137 — server-side twin of the multi-sheet header-guard data loss.
 *
 * `convertHeaderRowWritesToAddRow` filtered on `row === 0` without reading
 * `sheetName`, merged every match into one sheet-less ADD_ROW, and gated the
 * whole thing on the ACTIVE sheet's emptiness. In the reported incident the
 * active sheet happened to be empty, so this pass let the actions through and
 * the client-side copy did the damage — but on any workbook whose active sheet
 * has data, the same batch is destroyed here instead, before it is ever sent.
 */

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const HEADERS = [
  'Unit No', 'Guest', 'Guest Name', 'Check In', 'Check Out',
  'Rate Per Night', 'Total Amount', 'Source', 'Payment Status', 'Bank Account',
];

function monthlyLedgerBatch(): SheetActionPayload[] {
  const actions: SheetActionPayload[] = [];
  for (const month of MONTHS) {
    actions.push({ type: 'ADD_SHEET', name: month });
  }
  actions.push({ type: 'ADD_SHEET', name: 'Main' });
  for (const month of MONTHS) {
    HEADERS.forEach((header, col) => {
      actions.push({ type: 'SET_CELL', sheetName: month, row: 0, col, value: header });
    });
  }
  actions.push({ type: 'SET_CELL', sheetName: 'Main', row: 0, col: 0, value: 'Dashboard' });
  return actions;
}

/** A workbook whose only existing sheet has real headers — the dangerous case. */
const populatedContext: WorkbookContext = {
  activeSheet: 'Purchase Register',
  sheets: [
    {
      sheetName: 'Purchase Register',
      usedRange: 'A1:C10',
      rowCount: 10,
      colCount: 3,
      headers: ['Date', 'Supplier', 'Amount'],
      sampleData: [],
    },
  ],
};

describe('sheet-header-state util (TASKS.md #137)', () => {
  it('keys an action by its own sheet, falling back to the active sheet', () => {
    expect(sheetKeyOf({ type: 'SET_CELL', sheetName: 'January' }, 'Main')).toBe('january');
    expect(sheetKeyOf({ type: 'SET_CELL' }, 'Main')).toBe('main');
  });

  it('collects sheets the batch creates, including COPY_SHEET destinations', () => {
    expect(
      sheetsCreatedInBatch([
        { type: 'ADD_SHEET', name: 'January' },
        { type: 'CREATE_SHEET', sheetName: 'Main' },
        { type: 'COPY_SHEET', sheetName: 'Main', newName: 'Main Copy' },
      ]),
    ).toEqual(new Set(['january', 'main', 'main copy']));
  });

  it('treats every sheet created in the batch as having no header row', () => {
    const states = resolveSheetHeaderStates(monthlyLedgerBatch(), populatedContext, false);
    for (const month of MONTHS) {
      expect(states.get(month.toLowerCase())).toBe(false);
    }
    expect(states.get('main')).toBe(false);
  });

  it('still reports a real existing sheet as having a header row', () => {
    const states = resolveSheetHeaderStates(monthlyLedgerBatch(), populatedContext, false);
    expect(states.get('purchase register')).toBe(true);
  });

  it('groups a 13-sheet batch into 13 groups, never one', () => {
    const groups = groupActionsBySheet(monthlyLedgerBatch(), 'Purchase Register');
    // 12 months + Main. ADD_SHEET actions carry `name`, not `sheetName`, so they
    // key to the active sheet — that group exists too, and must stay separate.
    for (const month of MONTHS) {
      const group = groups.get(month.toLowerCase()) ?? [];
      expect(group.filter((a) => a.type === 'SET_CELL')).toHaveLength(HEADERS.length);
    }
    const main = groups.get('main') ?? [];
    expect(main.filter((a) => a.type === 'SET_CELL')).toHaveLength(1);
  });

  it('never pools row-1 writes from different sheets into one group', () => {
    const groups = groupActionsBySheet(monthlyLedgerBatch(), 'Purchase Register');
    const december = groups.get('december') ?? [];
    // "Dashboard" belongs to Main. In the incident it won column A on December.
    expect(december.some((a) => a.value === 'Dashboard')).toBe(false);
    expect(december.find((a) => a.col === 0)?.value).toBe('Unit No');
  });

  it('falls back to the active sheet state for an unknown sheet', () => {
    const states = resolveSheetHeaderStates(
      [{ type: 'SET_CELL', sheetName: 'Unknown', row: 0, col: 0, value: 'x' }],
      populatedContext,
      false,
    );
    // Not created here, not in context: the guard must stay conservative.
    expect(states.get('unknown')).toBeUndefined();
  });
});
