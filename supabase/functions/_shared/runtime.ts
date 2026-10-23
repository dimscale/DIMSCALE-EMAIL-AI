// Wires the real environment (Deno.env, fetch) for graph-webhook and auto-draft-cron. Not unit tested: tests inject fakes.
import { makeAppTokenProvider, readAppTokenCfg } from "./appToken.ts";
import { type Get, readSettings } from "./config.ts";
import { makeDb, readDbEnv } from "./db.ts";
import { makeGraph } from "./graph.ts";
import { log } from "./log.ts";
import type { ProcDeps } from "./processor.ts";

export function buildRuntime(fn: string) {
  const get: Get = (k) => Deno.env.get(k);
  const dbEnv = readDbEnv(get), tok = readAppTokenCfg(get);
  const webhookUrl = get("GRAPH_WEBHOOK_URL")?.trim(), clientState = get("GRAPH_CLIENT_STATE");
  if (!dbEnv || !tok || !webhookUrl || !clientState) return null;
  const db = makeDb(dbEnv.url, dbEnv.key);
  const provider = makeAppTokenProvider(tok); // one per instance: caches the token
  const graph = makeGraph({ getToken: provider.getToken, invalidateToken: provider.invalidate, webhookUrl, clientState });
  return {
    get,
    db,
    graph,
    /** Settings are read on every call, so a changed secret or env value needs no redeploy. */
    procDeps(): ProcDeps {
      const { settings, defaulted } = readSettings(get);
      if (defaulted.length) log({ fn, reason: "config_default", count: defaulted.length });
      return { fn, db, graph, settings, get };
    },
  };
}
