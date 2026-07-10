import {
  CellRef,
  FormulaFunctionName,
  ParsedFormula,
  RangeRef,
} from './formula.types';

export const FUNCTION_NAMES: FormulaFunctionName[] = [
  'SUMIFS',
  'SUMIF',
  'SUM',
  'AVERAGEIFS',
  'AVERAGEIF',
  'AVERAGE',
  'COUNTIFS',
  'COUNTIF',
  'COUNTA',
  'COUNT',
  'XLOOKUP',
  'VLOOKUP',
  'HLOOKUP',
  'INDEX',
  'MATCH',
  'IFERROR',
  'IFNA',
  'IFS',
  'IF',
  'AND',
  'OR',
  'NOT',
  'MAX',
  'MIN',
  'LARGE',
  'SMALL',
  'ROUNDUP',
  'ROUNDDOWN',
  'ROUND',
  'NPV',
  'IRR',
  'PMT',
  'FV',
  'PV',
  'TEXT',
  'LEFT',
  'RIGHT',
  'MID',
  'CONCAT',
  'LEN',
  'EOMONTH',
  'EDATE',
  'DATE',
  'YEAR',
  'MONTH',
  'DAY',
];

const CROSS_SHEET_RANGE =
  /(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/g;
const CROSS_SHEET_CELL = /(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)/g;
const LOCAL_RANGE = /\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/g;
const LOCAL_CELL = /\$?([A-Z]+)\$?(\d+)/g;

const AGGREGATION_FNS: FormulaFunctionName[] = [
  'SUM',
  'AVERAGE',
  'COUNT',
  'COUNTA',
  'SUMIF',
  'SUMIFS',
  'COUNTIF',
  'COUNTIFS',
];
const LOOKUP_FNS: FormulaFunctionName[] = ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH'];
const CONDITIONAL_FNS: FormulaFunctionName[] = ['IF', 'IFS', 'IFERROR', 'IFNA'];

export function parseFormula(formula: string): ParsedFormula {
  if (!formula.startsWith('=')) {
    return emptyParsed(formula);
  }

  const upper = formula.toUpperCase();
  const functions = FUNCTION_NAMES.filter((fn) => upper.includes(`${fn}(`));

  const crossSheetRefs: string[] = [];
  const rangeRefs: RangeRef[] = [];
  const cellRefs: CellRef[] = [];

  const noXSRange = formula.replace(CROSS_SHEET_RANGE, (match, q1, q2, sc, sr, ec, er) => {
    const sheet = q1 ?? q2;
    crossSheetRefs.push(sheet);
    rangeRefs.push({
      sheet,
      startCol: sc,
      startRow: Number.parseInt(sr, 10),
      endCol: ec,
      endRow: Number.parseInt(er, 10),
      isAbsoluteStart: match.includes('$'),
      isAbsoluteEnd: match.includes('$'),
      raw: match,
    });
    return '[[XS_RANGE]]';
  });

  const noXSCell = noXSRange.replace(CROSS_SHEET_CELL, (match, q1, q2, col, row) => {
    const sheet = q1 ?? q2;
    if (!crossSheetRefs.includes(sheet)) crossSheetRefs.push(sheet);
    cellRefs.push({
      sheet,
      column: col,
      row: Number.parseInt(row, 10),
      isAbsoluteCol: match.includes('$'),
      isAbsoluteRow: match.includes('$'),
      raw: match,
    });
    return '[[XS_CELL]]';
  });

  const noRange = noXSCell.replace(LOCAL_RANGE, (match, sc, sr, ec, er) => {
    rangeRefs.push({
      startCol: sc,
      startRow: Number.parseInt(sr, 10),
      endCol: ec,
      endRow: Number.parseInt(er, 10),
      isAbsoluteStart: match.includes('$'),
      isAbsoluteEnd: match.includes('$'),
      raw: match,
    });
    return '[[RANGE]]';
  });

  noRange.replace(LOCAL_CELL, (match, col, row) => {
    if (!Number.isNaN(Number.parseInt(row, 10)) && /^[A-Z]{1,3}$/.test(col)) {
      cellRefs.push({
        column: col,
        row: Number.parseInt(row, 10),
        isAbsoluteCol: match.startsWith('$'),
        isAbsoluteRow: match.includes(`$${row}`),
        raw: match,
      });
    }
    return match;
  });

  return {
    raw: formula,
    functions,
    cellRefs,
    rangeRefs,
    crossSheetRefs,
    isAggregation:
      functions.some((f) => AGGREGATION_FNS.includes(f)) && rangeRefs.length > 0,
    isLookup: functions.some((f) => LOOKUP_FNS.includes(f)),
    isConditional: functions.some((f) => CONDITIONAL_FNS.includes(f)),
  };
}

function emptyParsed(raw: string): ParsedFormula {
  return {
    raw,
    functions: [],
    cellRefs: [],
    rangeRefs: [],
    crossSheetRefs: [],
    isAggregation: false,
    isLookup: false,
    isConditional: false,
  };
}
