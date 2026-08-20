import {
  REVERSIBILITY_CATALOG,
  ALL_REVERSIBILITY_CATALOG_TYPES,
  computeIrreversibleActionTypes,
} from '../src/audit/reversibility-catalog';
import { SheetActionType } from '../src/excel-ai/types/sheet-actions.types';

describe('reversibility-catalog (TASKS.md #18)', () => {
  it('gives every irreversible entry a real, non-empty reason', () => {
    for (const type of ALL_REVERSIBILITY_CATALOG_TYPES) {
      const entry = REVERSIBILITY_CATALOG[type];
      if (!entry.reversible) {
        expect(entry.reason.trim().length).toBeGreaterThan(10);
      }
    }
  });

  it('marks every action type with a real, built inverse (TASKS.md #10/#12-#17) as reversible', () => {
    const shouldBeReversible: SheetActionType[] = [
      'SET_CELL',
      'SET_FORMULA',
      'ADD_SHEET',
      'CREATE_SHEET',
      'DELETE_SHEET',
      'INSERT_COLUMN',
      'DELETE_COLUMN',
      'INSERT_ROW',
      'DELETE_ROW',
      'CREATE_TABLE',
      'MERGE_CELLS',
      'UNMERGE_CELLS',
      'CREATE_CHART',
    ];
    for (const type of shouldBeReversible) {
      expect(REVERSIBILITY_CATALOG[type].reversible).toBe(true);
    }
  });

  it('does NOT mark RENAME_SHEET/COPY_SHEET/DEFINE_NAMED_RANGE reversible despite being shadow-simulated', () => {
    // These are the exact gap this catalog exists to catch — see the file's own
    // header comment for why `virtual-apply-catalog.ts`'s `simulated: true` is not
    // an equivalent signal.
    expect(REVERSIBILITY_CATALOG.RENAME_SHEET.reversible).toBe(false);
    expect(REVERSIBILITY_CATALOG.COPY_SHEET.reversible).toBe(false);
    expect(REVERSIBILITY_CATALOG.DEFINE_NAMED_RANGE.reversible).toBe(false);
  });

  it('marks the base CREATE_CHART entry reversible, and DELETE_CHART irreversible as a revert-only synthetic action (TASKS.md #15)', () => {
    expect(REVERSIBILITY_CATALOG.CREATE_CHART.reversible).toBe(true);
    expect(REVERSIBILITY_CATALOG.DELETE_CHART.reversible).toBe(false);
  });

  it('marks the base CONDITIONAL_FORMAT entry reversible, and DELETE_CONDITIONAL_FORMAT irreversible as a revert-only synthetic action (TASKS.md #40)', () => {
    expect(REVERSIBILITY_CATALOG.CONDITIONAL_FORMAT.reversible).toBe(true);
    expect(REVERSIBILITY_CATALOG.DELETE_CONDITIONAL_FORMAT.reversible).toBe(false);
  });

  it('has exactly the action types the backend action union declares — no more, no fewer', () => {
    expect(ALL_REVERSIBILITY_CATALOG_TYPES.length).toBe(
      Object.keys(REVERSIBILITY_CATALOG).length,
    );
  });

  describe('computeIrreversibleActionTypes', () => {
    it('returns empty for a batch of only fully-reversible actions', () => {
      expect(computeIrreversibleActionTypes(['SET_CELL', 'INSERT_COLUMN', 'DELETE_SHEET'])).toEqual([]);
    });

    it('flags the irreversible types present, deduplicated', () => {
      const result = computeIrreversibleActionTypes([
        'SET_CELL',
        'FREEZE_PANES',
        'RENAME_SHEET',
        'FREEZE_PANES',
      ]);
      expect(result.sort()).toEqual(['FREEZE_PANES', 'RENAME_SHEET']);
    });

    it('fails closed on an unknown/unlisted type rather than assuming it is safe', () => {
      expect(computeIrreversibleActionTypes(['SOME_FUTURE_ACTION_TYPE'])).toEqual([
        'SOME_FUTURE_ACTION_TYPE',
      ]);
    });

    it('treats a plain CONDITIONAL_FORMAT create (no existingRuleId) as reversible (TASKS.md #40)', () => {
      expect(
        computeIrreversibleActionTypes([{ type: 'CONDITIONAL_FORMAT' }]),
      ).toEqual([]);
    });

    it('flags a CONDITIONAL_FORMAT instance with existingRuleId as irreversible — a modify, not a create', () => {
      expect(
        computeIrreversibleActionTypes([
          { type: 'CONDITIONAL_FORMAT', existingRuleId: 'cf-1' },
        ]),
      ).toEqual(['CONDITIONAL_FORMAT']);
    });

    it('flags CONDITIONAL_FORMAT for the whole batch when even one instance in it is a modify', () => {
      const result = computeIrreversibleActionTypes([
        { type: 'CONDITIONAL_FORMAT' }, // create — revertible on its own
        { type: 'CONDITIONAL_FORMAT', existingRuleId: 'cf-1' }, // modify — not
      ]);
      expect(result).toEqual(['CONDITIONAL_FORMAT']);
    });

    it('still accepts plain type strings for backward compatibility', () => {
      expect(computeIrreversibleActionTypes(['CONDITIONAL_FORMAT'])).toEqual([]);
    });
  });
});
