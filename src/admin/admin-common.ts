import { Connection } from 'mongoose';
import { ObjectId } from 'mongodb';

/** Monthly list prices — mirrors Dashboard/src/lib/data/common.ts's PLAN_PRICE_INR (cellix_backend razorpay-checkout.service.ts is the source of truth). */
export const PLAN_PRICE_INR: Record<string, number> = { beta: 899, solo: 1299, firm: 5999 };

/** India time has no DST, so a fixed offset is exact — matches Dashboard/src/lib/data/common.ts. */
export const TZ = '+05:30';

export interface AdminUserSummary {
  id: string;
  name: string;
  email: string;
  image: string | null;
}

export function num(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

export function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function isValidObjectIdString(id: string): boolean {
  return /^[0-9a-f]{24}$/i.test(id) && ObjectId.isValid(id);
}

/** Looks up `user` docs by hex id string, silently skipping anything not shaped like one. */
export async function usersById(connection: Connection, ids: Iterable<string | null | undefined>): Promise<Map<string, AdminUserSummary>> {
  const map = new Map<string, AdminUserSummary>();
  const db = connection.db;
  const unique = [...new Set([...ids].filter((id): id is string => Boolean(id)))];
  const objectIds = unique.filter(isValidObjectIdString).map((id) => new ObjectId(id));
  if (!db || objectIds.length === 0) return map;
  const docs = await db
    .collection('user')
    .find({ _id: { $in: objectIds } }, { projection: { name: 1, email: 1, image: 1 } })
    .toArray();
  for (const doc of docs) {
    const id = String(doc._id);
    map.set(id, { id, name: doc.name ?? '', email: doc.email ?? '', image: doc.image ?? null });
  }
  return map;
}

export interface AdminOwnerSummary {
  key: string;
  label: string;
  email?: string;
  userId?: string;
  /** Paid through the public checkout with an email that maps to no signed-in account. */
  guest: boolean;
}

/** billingEntityId is a user id for signed-in checkouts and a bare email for guest checkouts — mirrors Dashboard/src/lib/data/common.ts's resolveOwners. */
export async function resolveOwners(connection: Connection, keys: Iterable<string>): Promise<Map<string, AdminOwnerSummary>> {
  const unique = [...new Set(keys)];
  const byId = await usersById(connection, unique);
  const emails = unique.filter((k) => k.includes('@'));
  const byEmail = new Map<string, AdminUserSummary>();
  const db = connection.db;
  if (emails.length && db) {
    const docs = await db
      .collection('user')
      .find({ email: { $in: emails } }, { projection: { name: 1, email: 1 } })
      .toArray();
    for (const doc of docs) {
      byEmail.set(String(doc.email).toLowerCase(), { id: String(doc._id), name: doc.name ?? '', email: doc.email, image: null });
    }
  }

  const owners = new Map<string, AdminOwnerSummary>();
  for (const key of unique) {
    const user = byId.get(key);
    if (user) {
      owners.set(key, { key, label: user.name || user.email, email: user.email, userId: user.id, guest: false });
      continue;
    }
    if (key.includes('@')) {
      const match = byEmail.get(key.toLowerCase());
      owners.set(key, { key, label: key, email: key, userId: match?.id, guest: true });
      continue;
    }
    owners.set(key, { key, label: `Unknown (${key.slice(0, 8)}…)`, guest: false });
  }
  return owners;
}

export interface SeriesPoint {
  start: string;
  value: number;
}

const TZ_OFFSET_MS = 330 * 60_000;

/** Every bucket in the range, zero-filled — mirrors Dashboard/src/lib/data/common.ts's fillBuckets. */
export function fillBuckets(
  bucket: 'hour' | 'day',
  from: Date,
  to: Date,
  rows: { _id: Date; value: number }[],
): SeriesPoint[] {
  const size = bucket === 'hour' ? 3600_000 : 86_400_000;
  const align = (t: number) => Math.floor((t + TZ_OFFSET_MS) / size) * size - TZ_OFFSET_MS;
  const values = new Map(rows.map((r) => [new Date(r._id).getTime(), r.value]));
  const points: SeriesPoint[] = [];
  for (let t = align(from.getTime()); t <= to.getTime(); t += size) {
    points.push({ start: new Date(t).toISOString(), value: values.get(t) ?? 0 });
  }
  return points;
}
