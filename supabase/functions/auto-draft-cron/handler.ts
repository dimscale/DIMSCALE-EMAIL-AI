import { autoDraftsEnabled, type Get } from "../_shared/config.ts";
import type { Db } from "../_shared/db.ts";
import { log } from "../_shared/log.ts";
import { type ProcDeps, processClaimed } from "../_shared/processor.ts";
import { ctEqual } from "../_shared/safe.ts";
import { catchUp, syncSubscriptions } from "../_shared/subscriptions.ts";

export type CronDeps = { get: Get; db: Db | null; procDeps: () => ProcDeps | null };
const FN = "auto-draft-cron";

const json = (body: unknown, status: number) => new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });

export async function handle(req: Request, d: CronDeps): Promise<Response> {
  const t0 = Date.now(), id = crypto.randomUUID();
  const finish = (res: Response, reason: string) => {
    log({ fn: FN, id, method: req.method, status: res.status, reason, ms: Date.now() - t0 });
    return res;
  };
  // Authentication first (AC-24): nothing is read or done without the shared secret.
  if (!(await ctEqual(req.headers.get("x-cron-secret"), d.get("CRON_SHARED_SECRET")))) {
    return finish(json({ ok: false, error: "unauthorized" }, 401), "unauthorized");
  }
  if (req.method !== "POST") return finish(json({ ok: false, error: "method_not_allowed" }, 405), "method");
  let op: unknown;
  try {
    op = (await req.json() as { op?: unknown }).op;
  } catch {
    op = undefined;
  }
  if (op !== "process" && op !== "renew" && op !== "retention") return finish(json({ ok: false, error: "bad_op" }, 400), "bad_op");

  try {
    if (op === "retention") { // database only, no Graph or OpenAI: allowed even with the kill switch off
      if (!d.db) return finish(json({ ok: false }, 503), "not_configured");
      await d.db.retention();
      return finish(json({ ok: true }, 200), "retention");
    }
    if (!autoDraftsEnabled(d.get)) return finish(json({ ok: true, skipped: "kill_switch" }, 200), "kill_switch");
    const p = d.procDeps();
    if (!p) return finish(json({ ok: false }, 503), "not_configured");
    // "renew" retries users that were out of scope; "process" does not (avoids a 403 every 2 minutes).
    await syncSubscriptions(p, op === "renew");
    if (op === "process") {
      for (const u of await p.db.cronUsers()) {
        if (!u.active || !u.subscriptionId) continue;
        try {
          await catchUp(p, u, p.settings.maxAgeHours);
        } catch {
          log({ fn: FN, id, user_id: u.id, reason: "catchup_error" }); // one user's failure must not stop the others or processClaimed
        }
      }
      await processClaimed(p);
    }
    return finish(json({ ok: true }, 200), op);
  } catch {
    return finish(json({ ok: false, error: "server_error" }, 500), "error");
  }
}
