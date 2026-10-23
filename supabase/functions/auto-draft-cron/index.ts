import { buildRuntime } from "../_shared/runtime.ts";
import { handle } from "./handler.ts";

const rt = buildRuntime("auto-draft-cron");

Deno.serve((req) => handle(req, { get: (k) => Deno.env.get(k), db: rt?.db ?? null, procDeps: () => rt?.procDeps() ?? null }));
