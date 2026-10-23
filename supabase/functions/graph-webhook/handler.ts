import { autoDraftsEnabled, type Get } from "../_shared/config.ts";
import type { Db } from "../_shared/db.ts";
import { log } from "../_shared/log.ts";
import { ctEqual } from "../_shared/safe.ts";

export type WebhookDeps = {
  get: Get;
  db: Db | null;
  waitUntil: (p: Promise<unknown>) => void;
  runProcess: () => Promise<unknown>;
  runLifecycle: (userId: string, subscriptionId: string, kind: string) => Promise<unknown>;
};

const isRec = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null;
const FN = "graph-webhook";
const MAX_BODY = 512 * 1024;

export async function handle(req: Request, d: WebhookDeps): Promise<Response> {
  const t0 = Date.now(), id = crypto.randomUUID();
  const finish = (res: Response, reason: string, count?: number) => {
    log({ fn: FN, id, method: req.method, status: res.status, reason, count, ms: Date.now() - t0 });
    return res;
  };

  // Validation handshake (AC-1): echo the decoded token as text/plain. Nothing else happens.
  const token = new URL(req.url).searchParams.get("validationToken");
  if (token !== null && req.method === "POST") {
    return finish(new Response(token, { status: 200, headers: { "Content-Type": "text/plain", "X-Content-Type-Options": "nosniff" } }), "validation");
  }
  if (req.method !== "POST") return finish(new Response(null, { status: 405 }), "method");

  const raw = await req.text();
  let body: unknown;
  try {
    if (raw.length > MAX_BODY) throw new Error("big");
    body = JSON.parse(raw);
  } catch {
    return finish(new Response(null, { status: 400 }), "bad_body");
  }
  const items = isRec(body) && Array.isArray(body.value) ? body.value.slice(0, 500) : [];

  // Kill switch (AC-9): acknowledge, insert nothing, call nothing.
  if (!autoDraftsEnabled(d.get)) return finish(new Response(null, { status: 202 }), "kill_switch");
  if (!d.db) return finish(new Response(null, { status: 503 }), "not_configured");

  const secret = d.get("GRAPH_CLIENT_STATE");
  const ids = new Map<string, Set<string>>();
  const lifecycle: { userId: string; subId: string; kind: string }[] = [];
  const tenant = d.get("ENTRA_TENANT_ID")?.trim();
  const users = new Map<string, string | null>(); // one lookup per distinct subscription per request
  try {
    for (const it of items) {
      if (!isRec(it) || !(await ctEqual(typeof it.clientState === "string" ? it.clientState : null, secret))) {
        log({ fn: FN, id, reason: "client_state_mismatch" });
        continue;
      }
      const subId = typeof it.subscriptionId === "string" ? it.subscriptionId : "";
      // Only our own tenant (fail closed when ENTRA_TENANT_ID is not configured). Tenant GUIDs compare case-insensitively.
      const ownTenant = !!tenant && typeof it.tenantId === "string" && it.tenantId.toLowerCase() === tenant.toLowerCase();
      let userId: string | null = null;
      if (subId && ownTenant) {
        if (!users.has(subId)) users.set(subId, await d.db.userForSubscription(subId));
        userId = users.get(subId) ?? null;
      }
      if (!userId) {
        log({ fn: FN, id, reason: ownTenant ? "unknown_subscription" : "tenant_mismatch" });
        continue;
      }
      if (typeof it.lifecycleEvent === "string") {
        lifecycle.push({ userId, subId, kind: it.lifecycleEvent });
      } else if (it.changeType === "created" && isRec(it.resourceData) && typeof it.resourceData.id === "string") {
        if (!ids.has(userId)) ids.set(userId, new Set());
        ids.get(userId)!.add(it.resourceData.id);
      }
    }
    let n = 0;
    for (const [userId, set] of ids) n += await d.db.enqueue(userId, [...set]); // duplicates collapse on the unique key
    if (ids.size || lifecycle.length) {
      d.waitUntil((async () => {
        for (const l of lifecycle) {
          try {
            await d.runLifecycle(l.userId, l.subId, l.kind);
          } catch {
            log({ fn: FN, id, reason: "lifecycle_error" }); // one failing event must not skip the others
          }
        }
        try {
          await d.runProcess();
        } catch {
          log({ fn: FN, id, reason: "background_error" });
        }
      })());
    }
    return finish(new Response(null, { status: 202 }), "accepted", n);
  } catch {
    // The insert (or lookup) failed: 503 makes Graph retry (AC-3).
    return finish(new Response(null, { status: 503 }), "db_error");
  }
}
