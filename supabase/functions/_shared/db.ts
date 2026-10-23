// Database access through the RPC functions of the F-013 migration, with the secret key
// (bypasses RLS; the tables have no policies). No error detail is propagated: it could carry data.

export type Status = "off" | "pending" | "active" | "not_in_scope" | "error";
export type SettingsView = { enabled: boolean; status: Status; contacts: string[] };
export type Claimed = { id: string; userId: string; messageId: string; attempts: number; createdAt: string };
export type UserCtx = { oid: string; mailSha256: string; enabled: boolean; contacts: string[] };
export type CronUser = {
  id: string;
  oid: string;
  status: Status;
  active: boolean; // enabled and at least one contact
  subscriptionId: string | null;
  expiresAt: string | null;
  lastRenewedAt: string | null;
  lastCatchupAt: string | null;
};

export interface Db {
  getSettings(oid: string): Promise<SettingsView>;
  putSettings(oid: string, enabled: boolean, mailSha256: string, contacts: string[]): Promise<SettingsView>;
  deleteSettings(oid: string): Promise<SettingsView>;
  userForSubscription(subId: string): Promise<string | null>;
  userCtx(userId: string): Promise<UserCtx | null>;
  cronUsers(): Promise<CronUser[]>;
  setSubscription(userId: string, subId: string, expiresAt: string): Promise<void>;
  deleteSubscription(userId: string): Promise<void>;
  catchupDone(userId: string, atIso: string): Promise<void>;
  setUserStatus(userId: string, status: Status): Promise<void>;
  enqueue(userId: string, messageIds: string[]): Promise<number>;
  claim(limit: number, userId?: string): Promise<Claimed[]>;
  finish(rowId: string, status: "drafted" | "skipped" | "failed", reason: string): Promise<void>;
  retry(rowId: string, reason: string, nextAttemptAtIso: string): Promise<void>;
  countsToday(userId: string): Promise<{ user: number; global: number }>;
  retention(): Promise<void>;
}

export function makeDb(url: string, key: string, fetchFn: typeof fetch = fetch): Db {
  async function rpc<T>(name: string, args: Record<string, unknown>): Promise<T> {
    const headers: Record<string, string> = { apikey: key, "Content-Type": "application/json" };
    if (key.startsWith("eyJ")) headers.Authorization = `Bearer ${key}`; // legacy JWT keys only; sb_secret_ keys go in apikey alone
    let res: Response;
    try {
      res = await fetchFn(`${url}/rest/v1/rpc/${name}`, { method: "POST", headers, body: JSON.stringify(args), signal: AbortSignal.timeout(15_000) });
    } catch {
      throw new Error("db_error");
    }
    if (!res.ok) throw new Error("db_error");
    const t = await res.text();
    return (t ? JSON.parse(t) : null) as T;
  }
  return {
    getSettings: (oid) => rpc("ad_settings_get", { p_oid: oid }),
    putSettings: (oid, enabled, sha, contacts) => rpc("ad_settings_put", { p_oid: oid, p_enabled: enabled, p_sha: sha, p_contacts: contacts }),
    deleteSettings: (oid) => rpc("ad_settings_delete", { p_oid: oid }),
    userForSubscription: (subId) => rpc("ad_user_for_subscription", { p_sub: subId }),
    userCtx: (userId) => rpc("ad_user_ctx", { p_user: userId }),
    cronUsers: () => rpc("ad_cron_users", {}),
    setSubscription: (userId, subId, expiresAt) => rpc("ad_sub_set", { p_user: userId, p_sub: subId, p_expires: expiresAt }),
    deleteSubscription: (userId) => rpc("ad_sub_delete", { p_user: userId }),
    catchupDone: (userId, at) => rpc("ad_catchup_done", { p_user: userId, p_at: at }),
    setUserStatus: (userId, status) => rpc("ad_user_status", { p_user: userId, p_status: status }),
    enqueue: (userId, ids) => rpc("ad_enqueue", { p_user: userId, p_ids: ids }),
    claim: (limit, userId) => rpc("ad_claim", { p_limit: limit, p_user: userId ?? null }),
    finish: (id, status, reason) => rpc("ad_finish", { p_id: id, p_status: status, p_reason: reason }),
    retry: (id, reason, next) => rpc("ad_retry", { p_id: id, p_reason: reason, p_next: next }),
    countsToday: (userId) => rpc("ad_counts_today", { p_user: userId }),
    retention: () => rpc("ad_retention", {}),
  };
}

export function readDbEnv(get: (k: string) => string | undefined): { url: string; key: string } | null {
  const url = get("SUPABASE_URL")?.trim();
  let key = get("SUPABASE_SERVICE_ROLE_KEY")?.trim();
  try {
    const keys = JSON.parse(get("SUPABASE_SECRET_KEYS") ?? "null") as Record<string, unknown> | null;
    if (typeof keys?.default === "string") key = keys.default;
  } catch { /* fall back to the legacy key */ }
  return url && key ? { url, key } : null;
}
