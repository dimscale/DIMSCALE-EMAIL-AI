import type { JWTVerifyGetKey } from "jose";
import type { Db } from "../_shared/db.ts";
import type { Config } from "../_shared/entra.ts";
import { log } from "../_shared/log.ts";
import { normalizeAddress, sha256Hex } from "../_shared/safe.ts";
import { validateStateful } from "../_shared/stateful.ts";

const FN = "auto-reply-settings";
const ADDRESS = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
const MAX_CONTACTS = 50;
const MAX_BODY = 32 * 1024;

type Parsed = { enabled: boolean; ownAddress: string; contacts: string[] };

/** Strict body validation (AC-37). Returns null when invalid. */
export function parsePut(body: unknown): Parsed | null {
  if (typeof body !== "object" || body === null || Array.isArray(body)) return null;
  const b = body as Record<string, unknown>;
  const keys = Object.keys(b).sort().join(",");
  if (keys !== "contacts,enabled,ownAddress" || typeof b.enabled !== "boolean" || typeof b.ownAddress !== "string" || !Array.isArray(b.contacts)) return null;
  const valid = (s: string) => s.length <= 254 && ADDRESS.test(s);
  const own = normalizeAddress(b.ownAddress);
  if (!valid(own) || b.contacts.length > MAX_CONTACTS) return null;
  const seen = new Set<string>();
  for (const c of b.contacts) {
    if (typeof c !== "string") return null;
    const n = normalizeAddress(c);
    if (!valid(n) || n === own || seen.has(n)) return null;
    seen.add(n);
  }
  return { enabled: b.enabled, ownAddress: own, contacts: [...seen] };
}

export async function handle(req: Request, cfg: Config | null, keySet: JWTVerifyGetKey | null, db: Db | null): Promise<Response> {
  const t0 = Date.now(), id = crypto.randomUUID();
  let reason = "ok", count: number | undefined;
  const origin = req.headers.get("Origin");
  const cors: Record<string, string> = { Vary: "Origin" };
  if (cfg && origin && cfg.allowedOrigins.includes(origin)) cors["Access-Control-Allow-Origin"] = origin;
  const json = (body: unknown, status: number) => new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json" } });

  const res = await (async (): Promise<Response> => {
    if (!cfg || !keySet || !db) {
      reason = "not_configured";
      return json({ ok: false, error: "server_error" }, 500);
    }
    if (req.method === "OPTIONS") {
      reason = "preflight";
      return new Response(null, {
        status: 204,
        headers: { ...cors, "Access-Control-Allow-Headers": "authorization, content-type", "Access-Control-Allow-Methods": "GET, PUT, DELETE, OPTIONS" },
      });
    }
    if (req.method !== "GET" && req.method !== "PUT" && req.method !== "DELETE") {
      reason = "method";
      return json({ ok: false, error: "method_not_allowed" }, 405);
    }
    const m = /^Bearer ([^\s]+)$/.exec(req.headers.get("Authorization") ?? "");
    const auth = m ? await validateStateful(m[1], keySet, cfg) : "no_bearer";
    if (typeof auth === "string") {
      reason = auth;
      return json({ ok: false, error: "unauthorized" }, 401);
    }
    if ("unavailable" in auth) {
      reason = "jwks_unavailable";
      return json({ ok: false, error: "server_error" }, 503);
    }
    // Every row is keyed by the validated token's oid, never by anything in the body.
    try {
      if (req.method === "GET") return json(await db.getSettings(auth.oid), 200);
      if (req.method === "DELETE") return json(await db.deleteSettings(auth.oid), 200);
      let parsed: Parsed | null = null;
      try {
        const text = await req.text();
        if (text.length <= MAX_BODY) parsed = parsePut(JSON.parse(text));
      } catch { /* invalid JSON */ }
      if (!parsed) {
        reason = "payload_invalid";
        return json({ ok: false, error: "payload_invalid" }, 422);
      }
      count = parsed.contacts.length;
      return json(await db.putSettings(auth.oid, parsed.enabled, await sha256Hex(parsed.ownAddress), parsed.contacts), 200);
    } catch {
      reason = "db_error";
      return json({ ok: false, error: "server_error" }, 503);
    }
  })();
  log({ fn: FN, id, method: req.method, status: res.status, reason, count, ms: Date.now() - t0 });
  return res;
}
