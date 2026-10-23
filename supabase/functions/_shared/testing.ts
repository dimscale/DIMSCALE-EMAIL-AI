// Test helpers only (never imported by index.ts files): an in-memory Db mirroring the SQL semantics.
import type { Claimed, CronUser, Db, SettingsView, Status, UserCtx } from "./db.ts";
import { MAX_ATTEMPTS } from "./processor.ts";

export type Row = { id: string; userId: string; messageId: string; status: string; reason: string | null; attempts: number; nextAttemptAt: number; lockedUntil: number; createdAt: number; updatedAt: number };

export function fakeDb(nowMs: () => number = Date.now) {
  let seq = 0;
  const users = new Map<string, { id: string; oid: string; sha: string; enabled: boolean; status: Status; contacts: string[] }>();
  const subs = new Map<string, { subId: string; expiresAt: string; lastRenewedAt: string; lastCatchupAt: string | null }>();
  const queue: Row[] = [];
  const calls: string[] = [];
  let failEnqueue = false;
  const view = (u?: { enabled: boolean; status: Status; contacts: string[] }): SettingsView => u ? { enabled: u.enabled, status: u.status, contacts: [...u.contacts].sort() } : { enabled: false, status: "off", contacts: [] };
  const byOid = (oid: string) => [...users.values()].find((u) => u.oid === oid);

  const db: Db = {
    getSettings: (oid) => Promise.resolve(view(byOid(oid))),
    putSettings: (oid, enabled, sha, contacts) => {
      calls.push("putSettings");
      let u = byOid(oid);
      if (!u) users.set(oid, u = { id: `u${++seq}`, oid, sha, enabled: false, status: "off", contacts: [] });
      u.sha = sha;
      u.contacts = [...contacts];
      u.status = !(enabled && contacts.length) ? "off" : u.enabled && u.status !== "off" ? u.status : "pending";
      u.enabled = enabled;
      return Promise.resolve(view(u));
    },
    deleteSettings: (oid) => {
      calls.push("deleteSettings");
      const u = byOid(oid);
      if (u) Object.assign(u, { enabled: false, status: "off", contacts: [] });
      return Promise.resolve(view(u));
    },
    userForSubscription: (s) => Promise.resolve([...subs.entries()].find(([, v]) => v.subId === s)?.[0] ?? null),
    userCtx: (id) => {
      const u = users.get(id) ?? [...users.values()].find((x) => x.id === id);
      return Promise.resolve(u ? { oid: u.oid, mailSha256: u.sha, enabled: u.enabled, contacts: [...u.contacts] } as UserCtx : null);
    },
    cronUsers: () =>
      Promise.resolve([...users.values()].flatMap((u): CronUser[] => {
        const s = subs.get(u.id);
        const active = u.enabled && u.contacts.length > 0;
        return s || active ? [{ id: u.id, oid: u.oid, status: u.status, active, subscriptionId: s?.subId ?? null, expiresAt: s?.expiresAt ?? null, lastRenewedAt: s?.lastRenewedAt ?? null, lastCatchupAt: s?.lastCatchupAt ?? null }] : [];
      })),
    setSubscription: (id, subId, expiresAt) => {
      subs.set(id, { subId, expiresAt, lastRenewedAt: new Date(nowMs()).toISOString(), lastCatchupAt: subs.get(id)?.lastCatchupAt ?? null });
      return Promise.resolve();
    },
    deleteSubscription: (id) => (subs.delete(id), Promise.resolve()),
    catchupDone: (id, at) => {
      const s = subs.get(id);
      if (s) s.lastCatchupAt = at;
      return Promise.resolve();
    },
    setUserStatus: (id, status) => {
      const u = [...users.values()].find((x) => x.id === id);
      if (u && (status !== "active" || u.enabled)) u.status = status; // mirrors ad_user_status
      return Promise.resolve();
    },
    enqueue: (userId, ids) => {
      calls.push("enqueue");
      if (failEnqueue) return Promise.reject(new Error("db_error"));
      let n = 0;
      for (const messageId of ids) {
        if (queue.some((r) => r.userId === userId && r.messageId === messageId)) continue;
        queue.push({ id: `q${++seq}`, userId, messageId, status: "pending", reason: null, attempts: 0, nextAttemptAt: nowMs(), lockedUntil: 0, createdAt: nowMs(), updatedAt: nowMs() });
        n++;
      }
      return Promise.resolve(n);
    },
    claim: (limit, userId) => {
      const t = nowMs();
      for (const r of queue) if (r.status === "processing" && r.lockedUntil < t && r.attempts >= MAX_ATTEMPTS && (!userId || r.userId === userId)) Object.assign(r, { status: "failed", reason: "lease_expired", lockedUntil: 0, updatedAt: t }); // mirrors ad_claim
      const out: Claimed[] = [];
      const eligible = (r: Row) => (r.status === "pending" && r.nextAttemptAt <= t) || (r.status === "processing" && r.lockedUntil < t);
      for (const r of queue) {
        if (out.length >= limit) break;
        if (userId && r.userId !== userId) continue;
        if (!eligible(r) || out.some((o) => o.userId === r.userId)) continue;
        if (queue.some((x) => x.userId === r.userId && x.status === "processing" && x.lockedUntil >= t)) continue;
        Object.assign(r, { status: "processing", lockedUntil: t + 300_000, attempts: r.attempts + 1, updatedAt: t });
        out.push({ id: r.id, userId: r.userId, messageId: r.messageId, attempts: r.attempts, createdAt: new Date(r.createdAt).toISOString() });
      }
      return Promise.resolve(out);
    },
    finish: (id, status, reason) => {
      const r = queue.find((x) => x.id === id && x.status === "processing");
      if (r) Object.assign(r, { status, reason, lockedUntil: 0, updatedAt: nowMs() });
      return Promise.resolve();
    },
    retry: (id, reason, next) => {
      const r = queue.find((x) => x.id === id && x.status === "processing");
      if (r) Object.assign(r, { status: "pending", reason, nextAttemptAt: Date.parse(next), lockedUntil: 0 });
      return Promise.resolve();
    },
    countsToday: (userId) => {
      const d = queue.filter((r) => r.status === "drafted" || (r.status === "failed" && (r.reason ?? "").startsWith("ai_"))); // reached the AI step
      return Promise.resolve({ user: d.filter((r) => r.userId === userId).length, global: d.length });
    },
    retention: () => (calls.push("retention"), Promise.resolve()),
  };
  return {
    db, users, subs, queue, calls,
    failEnqueue: (v: boolean) => { failEnqueue = v; },
    addUser(oid: string, contacts: string[], sha = "", enabled = true, status: Status = "active") {
      const u = { id: `u${++seq}`, oid, sha, enabled, status, contacts };
      users.set(oid, u);
      return u;
    },
  };
}
