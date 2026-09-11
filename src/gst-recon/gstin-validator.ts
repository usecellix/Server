import {
  GstinValidationResult,
  NormalizedRowWithClientGstinColumn,
  ReconType,
} from './types';

const GSTIN_REGEX =
  /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}[Z]{1}[0-9A-Z]{1}$/;

export function normalizeGstin(raw: string): string {
  return String(raw ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}

export function validateGstinFormat(raw: string): GstinValidationResult {
  const normalized = normalizeGstin(raw);
  if (!normalized) {
    return { ok: false, blockingError: 'Client GSTIN is required.' };
  }
  if (!GSTIN_REGEX.test(normalized)) {
    return {
      ok: false,
      blockingError: `"${raw}" is not a valid GSTIN format.`,
    };
  }
  return { ok: true, normalizedGstin: normalized };
}

/**
 * Hard gate: portal statement GSTIN must equal selected clientGstin.
 */
export function validatePortalGstinMatchesClient(
  clientGstin: string,
  portalStatementGstin: string,
): GstinValidationResult {
  const client = normalizeGstin(clientGstin);
  const portal = normalizeGstin(portalStatementGstin);
  if (!portal) {
    return {
      ok: false,
      blockingError:
        'Could not find a GSTIN on the portal statement sheet. Confirm the GSTR file includes the taxpayer GSTIN.',
    };
  }
  if (client !== portal) {
    return {
      ok: false,
      blockingError:
        `Selected client GSTIN (${client}) does not match the portal ` +
        `statement GSTIN (${portal}). Stopping before any matching runs. ` +
        `Confirm you uploaded the correct client's GSTR file, or correct ` +
        `the selected client GSTIN.`,
    };
  }
  return { ok: true, normalizedGstin: client };
}

/**
 * Flag books rows whose client-side GSTIN differs from run clientGstin.
 * Mismatched rows are surfaced, not silently dropped.
 */
export function flagCrossGstinRows(
  rows: NormalizedRowWithClientGstinColumn[],
  clientGstin: string,
): { clean: NormalizedRowWithClientGstinColumn[]; crossGstin: NormalizedRowWithClientGstinColumn[] } {
  const client = normalizeGstin(clientGstin);
  const clean: NormalizedRowWithClientGstinColumn[] = [];
  const crossGstin: NormalizedRowWithClientGstinColumn[] = [];
  for (const row of rows) {
    const rowClientGstin = row.clientSideGstin
      ? normalizeGstin(row.clientSideGstin)
      : null;
    if (rowClientGstin && rowClientGstin !== client) {
      crossGstin.push(row);
    } else {
      clean.push(row);
    }
  }
  return { clean, crossGstin };
}

/**
 * Extract portal statement GSTIN from sheet grid metadata or header cells.
 * Looks for common GST portal labels and first valid GSTIN near the top.
 */
export function extractPortalStatementGstin(
  grid: unknown[][],
  _reconType?: ReconType,
): string {
  if (!Array.isArray(grid) || !grid.length) return '';

  const labelHints = [
    'gstin',
    'gstin of recipient',
    'gstin of supplier',
    'gstin/uin of recipient',
    'gstin of the taxpayer',
    'taxpayer gstin',
    'gstin of registered person',
  ];

  const scanLimit = Math.min(grid.length, 40);
  for (let r = 0; r < scanLimit; r++) {
    const row = grid[r];
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] ?? '')
        .trim()
        .toLowerCase()
        .replace(/[\s_\-./]+/g, ' ');
      if (!cell) continue;
      const isLabel = labelHints.some((h) => cell === h || cell.includes(h));
      if (!isLabel) continue;
      // Value often in next cell or same cell after colon
      const sameCellGstin = String(row[c] ?? '').match(
        /[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]/i,
      );
      if (sameCellGstin) {
        const v = validateGstinFormat(sameCellGstin[0]);
        if (v.ok) return v.normalizedGstin!;
      }
      for (let nc = c + 1; nc < Math.min(c + 4, row.length); nc++) {
        const candidate = normalizeGstin(String(row[nc] ?? ''));
        const v = validateGstinFormat(candidate);
        if (v.ok) return v.normalizedGstin!;
      }
    }
  }

  // Fallback: first valid GSTIN in top rows that isn't in a data header column pattern
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const row = grid[r];
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      const candidate = normalizeGstin(String(cell ?? ''));
      const v = validateGstinFormat(candidate);
      if (v.ok) return v.normalizedGstin!;
    }
  }

  return '';
}

/**
 * Extract a proposed client GSTIN from books register when unambiguous.
 */
export function proposeClientGstinFromBooks(
  rows: Array<{ clientSideGstin?: string | null }>,
): string | undefined {
  const unique = new Set<string>();
  for (const row of rows) {
    if (!row.clientSideGstin) continue;
    const n = normalizeGstin(row.clientSideGstin);
    if (validateGstinFormat(n).ok) unique.add(n);
  }
  if (unique.size === 1) return [...unique][0];
  return undefined;
}
