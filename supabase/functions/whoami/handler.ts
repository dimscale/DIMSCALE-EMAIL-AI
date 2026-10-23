import { type JWTVerifyGetKey } from "jose";
import { type Config, jwksUrl, readConfig, validateToken } from "../_shared/entra.ts";

// The validation code moved to _shared/entra.ts (F-013/3b); re-exported so the whoami tests stay untouched.
export { type Config, jwksUrl, readConfig, validateToken };

function json(body: unknown, status: number, headers: Record<string, string>): Response {
  return new Response(JSON.stringify(body), { status, headers: { ...headers, "Content-Type": "application/json" } });
}

export async function handle(req: Request, cfg: Config | null, keySet: JWTVerifyGetKey | null): Promise<Response> {
  const start = Date.now();
  const id = crypto.randomUUID();
  let reason = "ok";
  let res: Response;
  const origin = req.headers.get("Origin");
  const cors: Record<string, string> = { Vary: "Origin" };
  if (cfg && origin && cfg.allowedOrigins.includes(origin)) cors["Access-Control-Allow-Origin"] = origin;

  if (!cfg || !keySet) {
    reason = "not_configured";
    res = json({ ok: false, error: "server_error" }, 500, cors);
  } else if (req.method === "OPTIONS") {
    reason = "preflight";
    res = new Response(null, {
      status: 204,
      headers: { ...cors, "Access-Control-Allow-Headers": "authorization", "Access-Control-Allow-Methods": "GET, OPTIONS" },
    });
  } else if (req.method !== "GET") {
    reason = "method";
    res = json({ ok: false, error: "method_not_allowed" }, 405, cors);
  } else {
    const m = /^Bearer ([^\s]+)$/.exec(req.headers.get("Authorization") ?? "");
    const result = m ? await validateToken(m[1], keySet, cfg) : "no_bearer";
    if (typeof result === "string") {
      reason = result;
      res = json({ ok: false, error: "unauthorized" }, 401, cors);
    } else if ("unavailable" in result) {
      reason = "jwks_unavailable";
      res = json({ ok: false, error: "server_error" }, 503, cors);
    } else {
      res = json({ ok: true, name: result.name }, 200, cors);
    }
  }
  // Only id, status, reason code and duration (AC-16).
  console.log(JSON.stringify({ id, status: res.status, reason, ms: Date.now() - start }));
  return res;
}
