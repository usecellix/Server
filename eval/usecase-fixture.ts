/**
 * The workbook every `cellix-basic-usecases.html` probe runs against.
 *
 * Shaped like a real Indian CA purchase register rather than clean synthetic
 * data, because most of the guide's Q&A and fix use cases only have a right
 * answer when the data is messy: blank and truncated GSTINs, duplicate invoice
 * numbers, a trailing-space supplier name (Tally export), one amount stored as
 * text with a ₹ prefix, real IF/ROUND formulas in the tax columns, a second
 * sheet to look up against, and one hidden sheet.
 *
 * Ground truth is computed from this data (see `expectedFacts`) so answers can
 * be checked instead of eyeballed — TASKS.md #223 was found exactly that way.
 */
export const HEADERS = [
  'Invoice No',
  'Invoice Date',
  'Supplier Name',
  'GSTIN',
  'Taxable Amount',
  'IGST',
  'CGST',
  'SGST',
  'Narration',
];

/** 01-04-2024 as an Excel serial — Office.js hands the add-in serials, not strings. */
const FIRST_DATE_SERIAL = 45383;

const SUPPLIERS: Array<[string, string]> = [
  ['Kerala Electricals Pvt Ltd', '32AABCK1234F1Z5'],
  ['Murugan Traders', '33AAGFM5678K1Z2'],
  ['ABC Traders', '32AADCA9012L1Z8'],
  ['Krishnamurthy & Co', '29AAKFK3456M1Z1'],
  ['Sai Steels ', '27AAHCS7890N1Z4'], // trailing space, on purpose
];

const AMOUNTS = [118500, 45000, 520000, 76000, 150000, 23000, 99000, 310000, 64000, 12500];

export const DATA_ROW_COUNT = 30;

export interface Fixture {
  /** Header row + data rows, exactly as the add-in sends `sheetData`. */
  sheetData: unknown[][];
  formulas: string[][];
  workbookContext: Record<string, unknown>;
  lastRow: number;
}

export function buildFixture(): Fixture {
  const values: unknown[][] = [HEADERS.slice()];
  const formulas: string[][] = [HEADERS.slice()];

  for (let i = 0; i < DATA_ROW_COUNT; i += 1) {
    const excelRow = i + 2;
    const [name, fullGstin] = SUPPLIERS[i % SUPPLIERS.length];
    const gstin = i % 9 === 4 ? '' : i === 11 ? '32AADCA9012' : fullGstin;
    const amountNumber = AMOUNTS[i % AMOUNTS.length];
    const amount: unknown = i === 7 ? '₹3,10,000' : amountNumber;
    const intraState = gstin.startsWith('32');
    const storedAsText = typeof amount === 'string';
    const igst = intraState ? 0 : storedAsText ? '#VALUE!' : Math.round(amountNumber * 0.18);
    const half = !intraState ? 0 : storedAsText ? '#VALUE!' : Math.round(amountNumber * 0.09);
    const invoiceNo =
      i === 20 ? 'INV/2024/005' : i === 25 ? 'INV/2024/010' : `INV/2024/${String(i + 1).padStart(3, '0')}`;
    const narration = i % 7 === 3 ? 'Credit Note against INV' : 'Purchase of goods';

    values.push([
      invoiceNo,
      FIRST_DATE_SERIAL + i * 3,
      name,
      gstin,
      amount,
      igst,
      half,
      half,
      narration,
    ]);
    formulas.push([
      invoiceNo,
      String(FIRST_DATE_SERIAL + i * 3),
      name,
      gstin,
      String(amount),
      `=IF(LEFT(D${excelRow},2)="32",0,ROUND(E${excelRow}*0.18,0))`,
      `=IF(LEFT(D${excelRow},2)="32",ROUND(E${excelRow}*0.09,0),0)`,
      `=IF(LEFT(D${excelRow},2)="32",ROUND(E${excelRow}*0.09,0),0)`,
      narration,
    ]);
  }

  const sheet = (name: string, headers: string[], rows: unknown[][], isHidden = false) => ({
    sheetName: name,
    usedRange: `A1:${String.fromCharCode(64 + headers.length)}${rows.length + 1}`,
    rowCount: rows.length + 1,
    colCount: headers.length,
    headers,
    sampleData: rows,
    isHidden,
  });

  return {
    sheetData: values,
    formulas,
    lastRow: values.length,
    workbookContext: {
      activeSheet: 'Purchase Register',
      sheets: [
        sheet('Purchase Register', HEADERS, values.slice(1)),
        sheet(
          'GSTR-2A',
          ['GSTIN', 'Invoice No', 'Taxable Value', 'IGST'],
          [
            ['32AABCK1234F1Z5', 'INV/2024/001', 118500, 0],
            ['33AAGFM5678K1Z2', 'INV/2024/002', 45000, 8100],
          ],
        ),
        sheet('Summary', ['Particulars', 'Amount'], [['Total Purchases', 4254000]]),
        sheet('Working', ['Note'], [['scratch']], true),
      ],
    },
  };
}

/** Ground truth, computed from the fixture — never hand-typed. */
export function expectedFacts() {
  const { sheetData } = buildFixture();
  const rows = sheetData.slice(1);
  const toNumber = (v: unknown) => (typeof v === 'number' ? v : Number(String(v).replace(/[₹,]/g, '')));

  const perSupplier = new Map<string, number>();
  let total = 0;
  let keralaTotal = 0;
  let keralaRows = 0;
  let blankGstins = 0;

  for (const row of rows) {
    const amount = toNumber(row[4]);
    const supplier = String(row[2]).trim();
    total += amount;
    perSupplier.set(supplier, (perSupplier.get(supplier) ?? 0) + amount);
    if (String(row[3]).startsWith('32')) {
      keralaTotal += amount;
      keralaRows += 1;
    }
    if (!String(row[3]).trim()) blankGstins += 1;
  }

  const invoiceNumbers = rows.map((r) => String(r[0]));
  const duplicates = invoiceNumbers.filter((id, i) => invoiceNumbers.indexOf(id) !== i);

  return {
    dataRows: rows.length,
    total,
    keralaTotal,
    keralaRows,
    blankGstins,
    duplicates: [...new Set(duplicates)],
    maxAmount: Math.max(...rows.map((r) => toNumber(r[4]))),
    topSupplier: [...perSupplier.entries()].sort((a, b) => b[1] - a[1])[0],
    firstDate: '01-04-2024',
    lastDate: '27-06-2024',
  };
}
