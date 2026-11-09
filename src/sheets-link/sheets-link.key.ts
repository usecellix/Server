import { createHash, randomBytes } from 'crypto';

/**
 * Google Sheets add-in authentication (pricing-page claim "Excel + Google Sheets add-in").
 *
 * The Sheets add-in runs in Google Apps Script, server-to-server, so it can never hold the
 * browser session cookie the Excel add-in uses. Instead a signed-in user creates a LINK KEY
 * once (GET /sheets-link/new-key in a browser where they are logged in to CELLIX) and pastes
 * it into the Sheets sidebar. The add-in sends it as `x-cellix-sheets-key`; AuthGuard turns it
 * back into that user's identity. Only the SHA-256 of the key is stored; keys expire and can
 * be revoked.
 */
export const SHEETS_KEY_HEADER = 'x-cellix-sheets-key';
export const SHEETS_KEY_PREFIX = 'cxs_';
export const SHEETS_KEY_TTL_DAYS = 90;

export interface LinkedUser {
  userId: string;
  email: string | null;
  name: string | null;
}

export const newSheetsKey = (): string => `${SHEETS_KEY_PREFIX}${randomBytes(24).toString('base64url')}`;
export const hashSheetsKey = (key: string): string => createHash('sha256').update(key.trim()).digest('hex');
export const looksLikeSheetsKey = (key: unknown): key is string =>
  typeof key === 'string' && key.startsWith(SHEETS_KEY_PREFIX) && key.length >= 30 && key.length <= 80;

/**
 * Static bridge so AuthGuard (constructed in many modules with only AppConfigService) can
 * resolve a key without every module importing the Mongo model. SheetsLinkService fills it in.
 */
export const SheetsKeyResolver: { resolve: ((keyHash: string) => Promise<LinkedUser | null>) | null } = {
  resolve: null,
};
