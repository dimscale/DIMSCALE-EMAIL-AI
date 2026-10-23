import type { CronUser, Db } from "./db.ts";
import type { Graph } from "./graph.ts";
import { log } from "./log.ts";

type D = { fn: string; db: Db; graph: Graph; nowMs?: () => number };
const RENEW_WITHIN_MS = 48 * 3_600_000;
const GUARD_MS = 10 * 60_000; // no second PATCH within 10 min of the last one

/** Creates the subscription. 403 means the mailbox is outside the RBAC scope (F-013/AC-40). */
export async function ensureSubscription(d: D, u: { id: string; oid: string }): Promise<void> {
  const r = await d.graph.createSubscription(u.oid);
  if (r.ok) {
    await d.db.setSubscription(u.id, r.data.id, r.data.expiresAt);
    await d.db.setUserStatus(u.id, "active");
  } else {
    await d.db.setUserStatus(u.id, r.status === 403 ? "not_in_scope" : "error");
    log({ fn: d.fn, user_id: u.id, reason: r.status === 403 ? "subscription_forbidden" : "subscription_failed" });
  }
}

/** Renewal run: create missing, PATCH those expiring within 48 h, DELETE for disabled users. */
export async function syncSubscriptions(d: D, retryNotInScope: boolean): Promise<void> {
  const now = d.nowMs ?? Date.now;
  for (const u of await d.db.cronUsers()) {
    try {
      if (!u.active) {
        if (u.subscriptionId) {
          const r = await d.graph.deleteSubscription(u.subscriptionId);
          if (r.ok) await d.db.deleteSubscription(u.id);
          else log({ fn: d.fn, user_id: u.id, reason: "subscription_delete_failed" });
        }
      } else if (!u.subscriptionId) {
        if (u.status !== "not_in_scope" || retryNotInScope) await ensureSubscription(d, u);
      } else if (u.expiresAt && Date.parse(u.expiresAt) - now() < RENEW_WITHIN_MS && !(u.lastRenewedAt && now() - Date.parse(u.lastRenewedAt) < GUARD_MS)) {
        await renew(d, u);
      }
    } catch {
      log({ fn: d.fn, user_id: u.id, reason: "sync_error" });
    }
  }
}

export async function renew(d: D, u: Pick<CronUser, "id" | "oid"> & { subscriptionId: string | null }): Promise<void> {
  if (!u.subscriptionId) return await ensureSubscription(d, u);
  const r = await d.graph.renewSubscription(u.subscriptionId);
  if (r.ok) return await d.db.setSubscription(u.id, u.subscriptionId, r.data.expiresAt);
  if (r.status === 404) { // gone on Graph's side: recreate
    await d.db.deleteSubscription(u.id);
    return await ensureSubscription(d, u);
  }
  log({ fn: d.fn, user_id: u.id, reason: "subscription_renew_failed" });
}

/** Enqueues every Inbox ID since max(last catch-up, now - maxAge). Idempotent through the unique key. */
export async function catchUp(d: D, u: { id: string; oid: string; lastCatchupAt: string | null }, maxAgeHours: number): Promise<void> {
  const now = (d.nowMs ?? Date.now)();
  const since = Math.max(u.lastCatchupAt ? Date.parse(u.lastCatchupAt) : 0, now - maxAgeHours * 3_600_000);
  const r = await d.graph.listInboxIds(u.oid, new Date(since).toISOString());
  if (!r.ok) return log({ fn: d.fn, user_id: u.id, reason: "catchup_failed" });
  if (r.data.length) await d.db.enqueue(u.id, r.data);
  await d.db.catchupDone(u.id, new Date(now).toISOString());
}

/** Lifecycle notifications (F-013/AC-5). reauthorizationRequired: one PATCH (no separate reauthorize call). */
export async function handleLifecycle(d: D, userId: string, subId: string, kind: string, maxAgeHours: number): Promise<void> {
  const u = await d.db.userCtx(userId);
  if (!u) return;
  if (kind === "reauthorizationRequired") {
    await renew(d, { id: userId, oid: u.oid, subscriptionId: subId });
  } else if (kind === "subscriptionRemoved") {
    await d.db.deleteSubscription(userId);
    if (u.enabled && u.contacts.length > 0) {
      await ensureSubscription(d, { id: userId, oid: u.oid });
      await catchUp(d, { id: userId, oid: u.oid, lastCatchupAt: null }, maxAgeHours);
    }
  } else if (kind === "missed") {
    await catchUp(d, { id: userId, oid: u.oid, lastCatchupAt: null }, maxAgeHours);
  }
}
