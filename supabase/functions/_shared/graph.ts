// Graph client for F-013. The ONLY calls it can make are listed in ALLOWED: reading messages,
// creating a reply draft, and subscription CRUD. Anything else throws before a request is built.
// There is deliberately no code path that sends, moves, copies or deletes mail.

const BASE = "https://graph.microsoft.com/v1.0";
const G = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
const U = `^/users/${G}`;
const ALLOWED: { method: string; path: RegExp }[] = [
  { method: "GET", path: new RegExp(`${U}/messages/[^/]+$`, "i") },
  { method: "GET", path: new RegExp(`${U}/mailFolders\\('inbox'\\)$`, "i") },
  { method: "GET", path: new RegExp(`${U}/mailFolders\\('inbox'\\)/messages$`, "i") },
  { method: "POST", path: new RegExp(`${U}/messages/[^/]+/createReply$`, "i") },
  { method: "GET", path: /^\/subscriptions$/ },
  { method: "POST", path: /^\/subscriptions$/ },
  { method: "PATCH", path: /^\/subscriptions\/[^/]+$/ },
  { method: "DELETE", path: /^\/subscriptions\/[^/]+$/ },
];

export type GResult<T> = { ok: true; data: T } | { ok: false; status: number; retryAfterSec?: number; reason?: "token_unavailable" | "network" };

export type Msg = {
  odataType: string;
  subject: string;
  fromAddress: string;
  fromName: string;
  toAddresses: string[];
  isDraft: boolean;
  inference: string;
  receivedDateTime: string;
  headers: Record<string, string>;
  body: string;
  parentFolderId: string;
};

export type GraphDeps = {
  getToken: () => Promise<string>;
  invalidateToken?: () => void;
  fetchFn?: typeof fetch;
  webhookUrl: string;
  clientState: string;
  nowMs?: () => number;
};

const isRec = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null;
const str = (v: unknown) => (typeof v === "string" ? v : "");
const addr = (v: unknown) => (isRec(v) && isRec(v.emailAddress) ? str(v.emailAddress.address) : "");

export function makeGraph(d: GraphDeps) {
  const f = d.fetchFn ?? fetch;
  const now = d.nowMs ?? Date.now;

  async function call(method: string, url: string, opts: { body?: unknown; prefer?: string[] } = {}): Promise<GResult<unknown>> {
    if (!url.startsWith(`${BASE}/`)) throw new Error("graph_url_not_allowed");
    const pathname = new URL(url).pathname; // normalised: dot segments and encodings resolved
    if (!pathname.startsWith("/v1.0/")) throw new Error("graph_url_not_allowed");
    const path = pathname.slice("/v1.0".length);
    if (!ALLOWED.some((a) => a.method === method && a.path.test(path))) throw new Error("graph_call_not_allowed");
    let token: string;
    try {
      token = await d.getToken();
    } catch {
      return { ok: false, status: 0, reason: "token_unavailable" };
    }
    const headers = new Headers({ Authorization: `Bearer ${token}` });
    for (const p of opts.prefer ?? []) headers.append("Prefer", p);
    if (opts.body !== undefined) headers.set("Content-Type", "application/json");
    let res: Response;
    try {
      res = await f(url, { method, headers, body: opts.body === undefined ? undefined : JSON.stringify(opts.body), signal: AbortSignal.timeout(20_000) });
    } catch {
      return { ok: false, status: 0, reason: "network" };
    }
    if (res.status === 401) d.invalidateToken?.();
    if (!res.ok) {
      const ra = Number(res.headers.get("Retry-After"));
      return { ok: false, status: res.status, retryAfterSec: Number.isFinite(ra) && ra > 0 ? ra : undefined };
    }
    if (res.status === 204) return { ok: true, data: null };
    try {
      return { ok: true, data: await res.json() };
    } catch {
      return { ok: false, status: 502 };
    }
  }
  const map = <T, R>(r: GResult<T>, fn: (d: T) => R): GResult<R> => (r.ok ? { ok: true, data: fn(r.data) } : r);

  const IMMUTABLE = 'IdType="ImmutableId"';
  const msgUrl = (oid: string, id: string) => `${BASE}/users/${oid}/messages/${encodeURIComponent(id)}`;
  const inbox = (oid: string) => `${BASE}/users/${oid}/mailFolders('inbox')`;
  const SELECT = "subject,from,toRecipients,isDraft,inferenceClassification,receivedDateTime,internetMessageHeaders,uniqueBody,parentFolderId";

  return {
    /** First read of every row: the sender only (F-013 filter 2b). */
    async readFrom(oid: string, id: string): Promise<GResult<{ from: string }>> {
      const r = await call("GET", `${msgUrl(oid, id)}?$select=from`, { prefer: [IMMUTABLE] });
      return map(r, (j) => ({ from: isRec(j) ? addr(j.from) : "" }));
    },

    /** Second read, only for allow-listed senders: exact minimal field list, text body, immutable IDs. */
    async readMessage(oid: string, id: string): Promise<GResult<Msg>> {
      const r = await call("GET", `${msgUrl(oid, id)}?$select=${SELECT}`, { prefer: ['outlook.body-content-type="text"', IMMUTABLE] });
      return map(r, (j): Msg => {
        const m = isRec(j) ? j : {};
        const from = isRec(m.from) ? m.from : {};
        const headers: Record<string, string> = {};
        if (Array.isArray(m.internetMessageHeaders)) {
          for (const h of m.internetMessageHeaders) {
            if (isRec(h) && typeof h.name === "string" && !(h.name.toLowerCase() in headers)) headers[h.name.toLowerCase()] = str(h.value);
          }
        }
        return {
          odataType: str(m["@odata.type"]),
          subject: str(m.subject),
          fromAddress: addr(from),
          fromName: isRec(from.emailAddress) ? str(from.emailAddress.name) : "",
          toAddresses: Array.isArray(m.toRecipients) ? m.toRecipients.map(addr).filter(Boolean) : [],
          isDraft: m.isDraft !== false, // fail closed
          inference: str(m.inferenceClassification),
          receivedDateTime: str(m.receivedDateTime),
          headers,
          body: isRec(m.uniqueBody) ? str(m.uniqueBody.content) : "",
          parentFolderId: str(m.parentFolderId),
        };
      });
    },

    /** Folder id only (no mail content), for filter 3. */
    async inboxFolderId(oid: string): Promise<GResult<string>> {
      return map(await call("GET", `${inbox(oid)}?$select=id`, { prefer: [IMMUTABLE] }), (j) => (isRec(j) ? str(j.id) : ""));
    },

    /** Catch-up: Inbox message IDs since a time, `$select=id` only, following @odata.nextLink (max 10 pages; `truncated` if more remain). */
    async listInboxIds(oid: string, sinceIso: string): Promise<GResult<{ ids: string[]; truncated: boolean }>> {
      // $orderby is allowed with this $filter because it uses the same property (https://learn.microsoft.com/graph/api/user-list-messages).
      let url: string | null = `${inbox(oid)}/messages?$filter=${encodeURIComponent(`receivedDateTime ge ${sinceIso}`)}&$orderby=receivedDateTime%20desc&$select=id&$top=50`;
      const ids: string[] = [];
      for (let page = 0; url && page < 10; page++) {
        let r: GResult<unknown>;
        try {
          r = await call("GET", url, { prefer: [IMMUTABLE] });
        } catch {
          return { ok: false, status: 502 }; // a nextLink outside the allow-list is a failure, not a crash
        }
        if (!r.ok) return r;
        const j = r.data;
        if (isRec(j) && Array.isArray(j.value)) for (const m of j.value) if (isRec(m) && typeof m.id === "string") ids.push(m.id);
        url = isRec(j) && typeof j["@odata.nextLink"] === "string" ? j["@odata.nextLink"] : null;
      }
      return { ok: true, data: { ids, truncated: url !== null } };
    },

    /** The one draft-creating call. Recipients are never set by us. */
    async createReply(oid: string, id: string, comment: string): Promise<GResult<{ id: string }>> {
      const r = await call("POST", `${msgUrl(oid, id)}/createReply`, {
        body: { comment, message: { categories: ["DIMSCALE AI draft"] } },
        prefer: [IMMUTABLE],
      });
      return map(r, (j) => ({ id: isRec(j) ? str(j.id) : "" }));
    },

    async createSubscription(oid: string): Promise<GResult<{ id: string; expiresAt: string }>> {
      const resource = `users/${oid}/mailFolders('inbox')/messages`;
      const expiry = new Date(now() + 10_000 * 60_000).toISOString();
      const r = await call("POST", `${BASE}/subscriptions`, {
        body: {
          changeType: "created",
          resource,
          notificationUrl: d.webhookUrl,
          lifecycleNotificationUrl: d.webhookUrl,
          clientState: d.clientState,
          expirationDateTime: expiry,
        },
        prefer: [IMMUTABLE],
      });
      if (r.ok) return map(r, (j) => ({ id: isRec(j) ? str(j.id) : "", expiresAt: isRec(j) && str(j.expirationDateTime) ? str(j.expirationDateTime) : expiry }));
      if (r.status !== 409) return r;
      // Duplicate: adopt the existing subscription.
      const l = await call("GET", `${BASE}/subscriptions`);
      if (!l.ok) return l;
      const hit = isRec(l.data) && Array.isArray(l.data.value)
        ? l.data.value.find((s) => isRec(s) && str(s.resource).toLowerCase() === resource.toLowerCase() && str(s.notificationUrl) === d.webhookUrl)
        : undefined;
      return isRec(hit) && str(hit.id) ? { ok: true, data: { id: str(hit.id), expiresAt: str(hit.expirationDateTime) || expiry } } : r;
    },

    async renewSubscription(subId: string): Promise<GResult<{ expiresAt: string }>> {
      const expiry = new Date(now() + 10_000 * 60_000).toISOString();
      const r = await call("PATCH", `${BASE}/subscriptions/${encodeURIComponent(subId)}`, { body: { expirationDateTime: expiry } });
      return map(r, (j) => ({ expiresAt: isRec(j) && str(j.expirationDateTime) ? str(j.expirationDateTime) : expiry }));
    },

    /** 404 counts as deleted. */
    async deleteSubscription(subId: string): Promise<GResult<null>> {
      const r = await call("DELETE", `${BASE}/subscriptions/${encodeURIComponent(subId)}`);
      return !r.ok && r.status === 404 ? { ok: true, data: null } : map(r, () => null);
    },
  };
}
export type Graph = ReturnType<typeof makeGraph>;
