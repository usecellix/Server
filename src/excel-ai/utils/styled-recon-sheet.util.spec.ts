import { buildReconSheetStylingActions, StyledSheetBuilder } from './styled-recon-sheet.util';

describe('StyledSheetBuilder', () => {
  it('pushRow pads to width and returns the 0-based row index', () => {
    const b = new StyledSheetBuilder(4);
    const idx0 = b.pushRow(['a']);
    const idx1 = b.pushRow(['b', 'c']);
    expect(idx0).toBe(0);
    expect(idx1).toBe(1);
    expect(b.rows[0]).toEqual(['a', '', '', '']);
    expect(b.rows[1]).toEqual(['b', 'c', '', '']);
  });

  it('pushSection records title/header/data row indices correctly', () => {
    const b = new StyledSheetBuilder(3);
    b.pushRow(['sheet title']);
    b.pushSection('Section A', ['H1', 'H2', 'H3'], [
      ['r1c1', 'r1c2', 'r1c3'],
      ['r2c1', 'r2c2', 'r2c3'],
    ]);
    // row 0 = sheet title, row 1 = section title, row 2 = header, rows 3-4 = data, row 5 = spacer
    expect(b.sections).toHaveLength(1);
    expect(b.sections[0]).toEqual({
      titleRow: 1,
      headerRow: 2,
      headers: ['H1', 'H2', 'H3'],
      dataRowRange: [3, 4],
    });
    expect(b.rows[1]).toEqual(['Section A', '', '']);
    expect(b.rows[2]).toEqual(['H1', 'H2', 'H3']);
    expect(b.rows[3]).toEqual(['r1c1', 'r1c2', 'r1c3']);
    expect(b.rows[4]).toEqual(['r2c1', 'r2c2', 'r2c3']);
    expect(b.rows[5]).toEqual(['', '', '']);
  });

  it('pushSection is a no-op (records nothing, writes nothing) when dataRows is empty', () => {
    const b = new StyledSheetBuilder(2);
    b.pushRow(['title']);
    b.pushSection('Empty Section', ['H1', 'H2'], []);
    expect(b.sections).toHaveLength(0);
    expect(b.rows).toHaveLength(1);
  });

  it('multiple sections accumulate correct, non-overlapping row ranges', () => {
    const b = new StyledSheetBuilder(2);
    b.pushRow(['title']); // row 0
    b.pushSection('A', ['H1', 'H2'], [['1', '2']]); // title=1, header=2, data=[3,3], spacer=4
    b.pushSection('B', ['H1', 'H2'], [['3', '4'], ['5', '6']]); // title=5, header=6, data=[7,8], spacer=9
    expect(b.sections[0].dataRowRange).toEqual([3, 3]);
    expect(b.sections[1].titleRow).toBe(5);
    expect(b.sections[1].headerRow).toBe(6);
    expect(b.sections[1].dataRowRange).toEqual([7, 8]);
  });

  it('beginSummaryBlock/endSummaryBlock records the row range of the metric rows in between', () => {
    const b = new StyledSheetBuilder(2);
    b.pushRow(['sheet title']); // row 0
    b.pushRow(['SUMMARY']); // row 1
    b.beginSummaryBlock();
    b.pushRow(['Matched (exact)', 5]); // row 2
    b.pushRow(['Blank GSTIN', 1]); // row 3
    b.pushRow(['Portal-only', 2]); // row 4
    b.endSummaryBlock();
    b.pushRow(['']); // row 5, not part of the block

    expect(b.summaryValueRange).toEqual([2, 4]);
  });

  it('summaryValueRange is null when beginSummaryBlock() was never called', () => {
    const b = new StyledSheetBuilder(2);
    b.pushRow(['title']);
    expect(b.summaryValueRange).toBeNull();
  });

  it('endSummaryBlock() without a matching beginSummaryBlock() is a safe no-op', () => {
    const b = new StyledSheetBuilder(2);
    b.pushRow(['title']);
    b.endSummaryBlock();
    expect(b.summaryValueRange).toBeNull();
  });
});

describe('buildReconSheetStylingActions', () => {
  it('formats the sheet title row and every section title + header row with navy fill and white bold', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 3,
      sheetTitleRow: 0,
      sections: [
        { titleRow: 1, headerRow: 2, headers: ['GSTIN', 'Vendor Name', 'Taxable Value'], dataRowRange: [3, 4] },
      ],
    });

    const formatRanges = actions.filter((a) => a.type === 'FORMAT_RANGE');
    // sheet title + section title + section header = 3 navy bands, plus 1 currency number-format range = 4
    const navyBands = formatRanges.filter((a) => a.format?.fillColor === '#203764');
    expect(navyBands).toHaveLength(3);
    for (const band of navyBands) {
      expect(band.format?.bold).toBe(true);
      expect(band.format?.fontColor).toBe('#FFFFFF');
      expect(band.sheetName).toBe('Sheet1');
    }
    expect(navyBands[0].range).toBe('A1:C1'); // sheet title
    expect(navyBands[1].range).toBe('A2:C2'); // section title
    expect(navyBands[2].range).toBe('A3:C3'); // section header
    expect(navyBands[2].format?.horizontalAlignment).toBe('center');
  });

  it('applies currency number format to Taxable Value / IGST / CGST / SGST / Tax Amount columns only', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 5,
      sheetTitleRow: 0,
      sections: [
        {
          titleRow: 1,
          headerRow: 2,
          headers: ['GSTIN', 'Taxable Value', 'CGST', 'Document Type', 'Reason'],
          dataRowRange: [3, 5],
        },
      ],
    });

    const currencyRanges = actions.filter(
      (a) => a.type === 'FORMAT_RANGE' && a.format?.numberFormat === '#,##0.00',
    );
    expect(currencyRanges).toHaveLength(2);
    expect(currencyRanges.map((a) => a.range).sort()).toEqual(['B4:B6', 'C4:C6']);
  });

  it('applies date number format to Invoice Date columns', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 3,
      sheetTitleRow: 0,
      sections: [
        { titleRow: 1, headerRow: 2, headers: ['GSTIN', 'Invoice Date', 'Vendor Name'], dataRowRange: [3, 3] },
      ],
    });

    const dateRanges = actions.filter(
      (a) => a.type === 'FORMAT_RANGE' && a.format?.numberFormat === 'dd-mmm-yyyy',
    );
    expect(dateRanges).toHaveLength(1);
    expect(dateRanges[0].range).toBe('B4:B4');
  });

  it('skips number-format and banding actions for a section with no data rows', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 2,
      sheetTitleRow: 0,
      sections: [{ titleRow: 1, headerRow: 2, headers: ['GSTIN', 'Taxable Value'], dataRowRange: null }],
    });
    expect(actions.some((a) => a.type === 'CONDITIONAL_FORMAT')).toBe(false);
    expect(actions.some((a) => a.format?.numberFormat)).toBe(false);
  });

  it('adds row banding (CONDITIONAL_FORMAT) covering exactly the data row range for each section', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 4,
      sheetTitleRow: 0,
      sections: [{ titleRow: 1, headerRow: 2, headers: ['A', 'B', 'C', 'D'], dataRowRange: [3, 10] }],
    });
    const banding = actions.filter((a) => a.type === 'CONDITIONAL_FORMAT');
    expect(banding).toHaveLength(1);
    expect(banding[0].range).toBe('A4:D11');
    expect(banding[0].rule).toEqual({
      kind: 'formula',
      formula: '=MOD(ROW(),2)=0',
      format: { fillColor: '#F2F2F2' },
    });
  });

  it('sets a wide column width for Vendor Name / Explanation and a narrow width for GSTIN / Source Row', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 3,
      sheetTitleRow: 0,
      sections: [
        { titleRow: 1, headerRow: 2, headers: ['GSTIN', 'Vendor Name', 'Explanation'], dataRowRange: [3, 3] },
      ],
    });
    const widths = actions.filter((a) => a.type === 'SET_COLUMN_WIDTH');
    expect(widths).toHaveLength(3);
    const byCol = new Map(widths.map((a) => [a.col, a.width]));
    expect(byCol.get(0)).toBe(110); // GSTIN — narrow
    expect(byCol.get(1)).toBe(320); // Vendor Name — wide
    expect(byCol.get(2)).toBe(320); // Explanation — wide
  });

  it('when the same column index is "wide" in one section and "narrow" in another, the wider width wins', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 1,
      sheetTitleRow: 0,
      sections: [
        { titleRow: 1, headerRow: 2, headers: ['GSTIN'], dataRowRange: [3, 3] },
        { titleRow: 5, headerRow: 6, headers: ['Vendor Name'], dataRowRange: [7, 7] },
      ],
    });
    const width = actions.find((a) => a.type === 'SET_COLUMN_WIDTH' && a.col === 0);
    expect(width?.width).toBe(320);
  });

  it('freezes panes at the row immediately below the sheet title', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 2,
      sheetTitleRow: 0,
      sections: [],
    });
    const freeze = actions.find((a) => a.type === 'FREEZE_PANES');
    expect(freeze).toEqual({ type: 'FREEZE_PANES', sheetName: 'Sheet1', freezeRows: 1 });
  });

  it('centers the SUMMARY block value column (B) when summaryValueRange is set — alignment only, e.g. B5:B18', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Missed vs GSTR-2B',
      columnCount: 13,
      sheetTitleRow: 0,
      sections: [],
      summaryValueRange: [4, 17], // 0-based rows 4..17 -> Excel rows 5..18
    });
    const summaryFormats = actions.filter(
      (a) => a.type === 'FORMAT_RANGE' && a.range === 'B5:B18',
    );
    expect(summaryFormats).toHaveLength(1);
    expect(summaryFormats[0].format).toEqual({ horizontalAlignment: 'center' });
    // Alignment only — no number format, wrap text, or width/height side effects from this action.
    expect(summaryFormats[0].format?.numberFormat).toBeUndefined();
    expect(summaryFormats[0].format?.wrapText).toBeUndefined();
  });

  it('emits no SUMMARY value-column formatting when summaryValueRange is absent', () => {
    const actions = buildReconSheetStylingActions({
      sheetName: 'Sheet1',
      columnCount: 2,
      sheetTitleRow: 0,
      sections: [],
    });
    expect(actions.some((a) => a.type === 'FORMAT_RANGE' && a.range?.startsWith('B'))).toBe(false);
  });
});
