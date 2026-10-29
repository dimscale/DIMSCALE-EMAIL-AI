import type { JWTVerifyGetKey } from "jose";
import { type AiConfig, type AiFail, type DraftInput, generateDraft } from "../_shared/ai.ts";
import type { Db } from "../_shared/db.ts";
import type { Config } from "../_shared/entra.ts";
import { log } from "../_shared/log.ts";
import { sha256Hex } from "../_shared/safe.ts";
import { validateStateful } from "../_shared/stateful.ts";
import { generateSummary, hasLink } from "./summary.ts";

const FN = "analyze-message";
const MAX_RAW = 16 * 1024;
const MAX_CAP = 1_000_000;
// shortcut: effort and token cap are constants, not env (spec decision 9); make them env only if the user asks.
const EFFORT = "low", MAX_OUTPUT_TOKENS = 2000;
const NO_LINKS: AiFail = { ok: false, reason: "ai_output_invalid", retryable: false };

/** Body as text, or null when over `max` bytes: Content-Length is checked first, then a byte counter stops an unlabelled stream. */
async function readCapped(req: Request, max: number): Promise<string | null> {
  if (Number(req.headers.get("Content-Length") ?? 0) > max) return null;
  const reader = req.body?.getReader(), dec = new TextDecoder();
  let text = "", total = 0;
  while (reader) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > max) {
      await reader.cancel();
      return null;
    }
    text += dec.decode(value, { stream: true });
  }
  return text + dec.decode();
}

type Payload = DraftInput & { action: "summarize" | "reply" };

/** Strict payload validation (AC-3). Returns null when invalid. Lengths are UTF-16 units, like the pane. */
export function parsePayload(raw: unknown): Payload | null {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) return null;
  const b = raw as Record<string, unknown>;
  if (Object.keys(b).sort().join(",") !== "action,body,fromDomain,fromName,subject") return null;
  const { action, body, fromDomain, fromName, subject } = b;
  if (action !== "summarize" && action !== "reply") return null;
  if (typeof body !== "string" || typeof fromDomain !== "string" || typeof fromName !== "string" || typeof subject !== "string") return null;
  if (subject.length > 200 || fromName.length > 100 || fromDomain.length > 253 || !/^[^\s@<>]*$/.test(fromDomain)) return null;
  if (body.length > 4000 || body.trim() === "") return null;
  return { action, body, fromDomain, fromName, subject };
}

const posInt = (raw: string | undefined, dflt: number) => (raw !== undefined && /^\d+$/.test(raw.trim()) && Number(raw) > 0 && Number(raw) <= MAX_CAP ? Number(raw) : dflt);

export async function handle(
  req: Request, cfg: Config | null, keySet: JWTVerifyGetKey | null, db: Db | null,
  env: (k: string) => string | undefined = () => undefined, fetchFn: typeof fetch = fetch,
): Promise<Response> {
  const t0 = Date.now(), id = crypto.randomUUID();
  let reason = "ok";
  const origin = req.headers.get("Origin");
  const cors: Record<string, string> = { Vary: "Origin" };
  if (cfg && origin && cfg.allowedOrigins.includes(origin)) cors["Access-Control-Allow-Origin"] = origin;
  const json = (body: unknown, status: number) => new Response(JSON.stringify(body), { status, headers: { ...cors, "Content-Type": "application/json" } });
  const aiFail = (r: AiFail) => {
    reason = r.reason;
    return r.reason === "ai_unavailable" ? json({ ok: false, error: "ai_unavailable", retryable: r.retryable }, 503) : json({ ok: false, error: "ai_output_invalid" }, 502);
  };

  const res = await (async (): Promise<Response> => {
    if (!cfg || !keySet || !db) {
      reason = "not_configured";
      return json({ ok: false, error: "server_error" }, 500);
    }
    if (req.method === "OPTIONS") {
      reason = "preflight";
      return new Response(null, { status: 204, headers: { ...cors, "Access-Control-Allow-Headers": "authorization, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" } });
    }
    if (req.method !== "POST") {
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
    let payload: Payload | null = null;
    try {
      const text = await readCapped(req, MAX_RAW);
      if (text !== null) payload = parsePayload(JSON.parse(text));
    } catch { /* invalid JSON or unreadable body */ }
    if (!payload) {
      reason = "payload_invalid";
      return json({ ok: false, error: "payload_invalid" }, 422);
    }
    if (env("ANALYZE_DISABLED")?.trim() === "true") {
      reason = "disabled";
      return json({ ok: false, error: "disabled" }, 503);
    }
    const apiKey = env("OPENAI_API_KEY")?.trim() ?? "", model = env("OPENAI_MODEL")?.trim() ?? "";
    if (!apiKey || !model) {
      reason = "ai_not_configured";
      return json({ ok: false, error: "ai_not_configured" }, 503);
    }
    let slot: boolean;
    try {
      slot = await db.analyzeTake(await sha256Hex(auth.oid), posInt(env("ANALYZE_DAILY_CAP_PER_USER"), 50), posInt(env("ANALYZE_DAILY_CAP_GLOBAL"), 1500));
    } catch {
      reason = "db_error";
      return json({ ok: false, error: "server_error" }, 503);
    }
    if (!slot) {
      reason = "cap_reached";
      return json({ ok: false, error: "cap_reached" }, 429);
    }
    const ai: AiConfig = { apiKey, model, maxOutputTokens: MAX_OUTPUT_TOKENS, reasoningEffort: EFFORT };
    const { action, ...input } = payload;
    if (action === "summarize") {
      const r = await generateSummary(input, ai, fetchFn);
      return r.ok ? json({ ok: true, action, summary: r.summary }, 200) : aiFail(r);
    }
    const r = await generateDraft(input, ai, fetchFn);
    if (r.ok && hasLink(r.reply)) return aiFail(NO_LINKS);
    return r.ok ? json({ ok: true, action, reply: r.reply }, 200) : aiFail(r);
  })();
  log({ fn: FN, id, method: req.method, status: res.status, reason, ms: Date.now() - t0 });
  return res;
}
