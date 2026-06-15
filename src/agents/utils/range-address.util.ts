function colIndexToLetter(index: number): string {
  let n = index + 1;
  let letter = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

/** Strip workbook/sheet qualifier — e.g. `'Purchases'!A1:M339` → `A1:M339`. */
export function stripSheetPrefix(address: string): string {
  const trimmed = address.trim().replace(/\$/g, '');
  const bang = trimmed.lastIndexOf('!');
  return bang >= 0 ? trimmed.slice(bang + 1) : trimmed;
}

const LOCAL_RANGE = /^[A-Za-z]+\d+(?::[A-Za-z]+\d+)?$/;

export function isLocalRangeAddress(address: string): boolean {
  return LOCAL_RANGE.test(stripSheetPrefix(address));
}

export function normalizeSortRangeAddress(params: {
  usedRange?: string;
  rowCount: number;
  columnCount: number;
  key: number;
}): string {
  const local = stripSheetPrefix(params.usedRange ?? '');
  if (local.includes(':') && isLocalRangeAddress(local)) {
    return local;
  }
  const endCol = colIndexToLetter(Math.max(params.columnCount - 1, params.key));
  return `A1:${endCol}${Math.max(params.rowCount, 2)}`;
}
