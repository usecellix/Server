import {
  assertDomainToolProvenance,
  buildWorkbookSourceRefsFromActions,
} from '../src/audit/utils/provenance.util';
import { extractFormulaPrecedents } from '../src/audit/utils/formula-precedents.util';
import { ChangeSetService } from '../src/audit/change-set.service';
import { Action, WorkbookContext } from '../src/agents/types/agent.types';
import { CellChange } from '../src/audit/types/change-set.types';

describe('formula precedents', () => {
  it('extracts sheet-qualified ranges from formulas', () => {
    expect(extractFormulaPrecedents('=SUM(Sheet2!C4:C40)', 'Sheet1')).toEqual([
      'Sheet2!C4:C40',
    ]);
  });

  it('uses active sheet when formula has unqualified refs', () => {
    expect(extractFormulaPrecedents('=B2*0.18', 'Invoices')).toEqual(['Invoices!B2']);
  });
});

describe('buildWorkbookSourceRefsFromActions', () => {
  it('builds workbook SourceRefs for Tier 2 formula actions', () => {
    const actions: Action[] = [
      {
        type: 'SET_FORMULA',
        sheetName: 'Invoices',
        row: 1,
        col: 2,
        formula: '=B2*0.18',
      },
    ];
    const refs = buildWorkbookSourceRefsFromActions(actions, 'Invoices', 'Invoices');
    expect(refs.length).toBeGreaterThan(0);
    expect(refs[0]).toEqual(
      expect.objectContaining({
        documentType: 'workbook',
        documentId: 'Invoices',
        rowOrLine: 'Invoices!B2',
      }),
    );
  });
});

describe('assertDomainToolProvenance', () => {
  it('allows non-domain-tool previews without sourceRefs', () => {
    expect(() => assertDomainToolProvenance(undefined)).not.toThrow();
    expect(() => assertDomainToolProvenance({ fromDomainTool: false })).not.toThrow();
  });

  it('fails when a domain-tool write reaches createPreview without sourceRefs', () => {
    expect(() =>
      assertDomainToolProvenance({ fromDomainTool: true, sourceRefs: [] }),
    ).toThrow(/non-empty sourceRefs/i);
  });
});

describe('ChangeSetService.createPreview provenance', () => {
  const context: WorkbookContext = {
    activeSheetName: 'Sheet1',
    sheets: [
      {
        name: 'Sheet1',
        usedRange: 'A1:C3',
        rowCount: 3,
        columnCount: 3,
        values: [
          ['A', 'B', 'C'],
          [1, 2, 3],
          [4, 5, 6],
        ],
        formulas: [
          ['', '', ''],
          ['', '', ''],
          ['', '', ''],
        ],
        numberFormats: [
          ['General', 'General', 'General'],
          ['General', 'General', 'General'],
          ['General', 'General', 'General'],
        ],
        structure: 'data_table',
      },
    ],
    namedRanges: [],
    tables: [],
  };

  it('attaches sourceRefs and exceptionFlags to CellChange entries', async () => {
    const created: { changes?: CellChange[] } = {};
    const model = {
      create: jest.fn(async (doc: { changes: CellChange[] }) => {
        created.changes = doc.changes;
        return {
          ...doc,
          status: 'previewed',
          timestamp: new Date(),
        };
      }),
    };

    const service = new ChangeSetService(model as never);
    const actions: Action[] = [
      { type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2*C2' },
    ];

    await service.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'calculate product',
      context,
      actions,
      provenance: {
        sourceRefs: [
          {
            documentType: 'workbook',
            documentId: 'Sheet1',
            rowOrLine: 'Sheet1!B2',
          },
        ],
        exceptionFlags: [
          {
            code: 'GST_NAME_FUZZY_MATCH',
            severity: 'flag',
            message: 'Vendor name matched fuzzily',
            affectedRows: [2],
          },
        ],
      },
    });

    expect(created.changes?.[0]?.sourceRefs?.[0]?.rowOrLine).toBe('Sheet1!B2');
    expect(created.changes?.[0]?.exceptionFlags?.[0]?.code).toBe('GST_NAME_FUZZY_MATCH');
  });

  it('rejects domain-tool createPreview without sourceRefs', async () => {
    const service = new ChangeSetService({ create: jest.fn() } as never);
    await expect(
      service.createPreview({
        conversationId: 'conv-1',
        traceId: 'trace-1',
        prompt: 'gst match',
        context,
        actions: [{ type: 'SET_FORMULA', sheetName: 'Sheet1', row: 1, col: 2, formula: '=B2' }],
        provenance: { fromDomainTool: true, sourceRefs: [] },
      }),
    ).rejects.toThrow(/non-empty sourceRefs/i);
  });

  it('leaves legacy changes without sourceRefs when provenance omitted', async () => {
    const created: { changes?: CellChange[] } = {};
    const model = {
      create: jest.fn(async (doc: { changes: CellChange[] }) => {
        created.changes = doc.changes;
        return { ...doc, status: 'previewed', timestamp: new Date() };
      }),
    };
    const service = new ChangeSetService(model as never);
    await service.createPreview({
      conversationId: 'conv-1',
      traceId: 'trace-1',
      prompt: 'bold A1',
      context,
      actions: [
        {
          type: 'SET_CELL',
          sheetName: 'Sheet1',
          row: 0,
          col: 0,
          value: 'x',
        },
      ],
    });

    // Tier 0/1 style writes without provenance stay backward-compatible
    expect(created.changes?.every((c) => c.sourceRefs === undefined)).toBe(true);
  });
});
