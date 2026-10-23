import { buildRuntime } from "../_shared/runtime.ts";
import { handleLifecycle } from "../_shared/subscriptions.ts";
import { processClaimed } from "../_shared/processor.ts";
import { handle } from "./handler.ts";

declare const EdgeRuntime: { waitUntil(p: Promise<unknown>): void };

const rt = buildRuntime("graph-webhook");

Deno.serve((req) =>
  handle(req, {
    get: (k) => Deno.env.get(k),
    db: rt?.db ?? null,
    waitUntil: (p) => EdgeRuntime.waitUntil(p),
    runProcess: () => processClaimed(rt!.procDeps()),
    runLifecycle: (userId, subId, kind) => {
      const d = rt!.procDeps();
      return handleLifecycle(d, userId, subId, kind, d.settings.maxAgeHours);
    },
  })
);
