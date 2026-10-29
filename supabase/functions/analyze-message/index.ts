import { createRemoteJWKSet } from "jose";
import { makeDb, readDbEnv } from "../_shared/db.ts";
import { jwksUrl, readConfig } from "../_shared/entra.ts";
import { handle } from "./handler.ts";

const get = (k: string) => Deno.env.get(k);
const cfg = readConfig(get);
const keySet = cfg ? createRemoteJWKSet(new URL(jwksUrl(cfg.tenantId))) : null;
const dbEnv = readDbEnv(get);
const db = dbEnv ? makeDb(dbEnv.url, dbEnv.key) : null;

// Env values (kill switch, caps, OpenAI key and model) are read per request: a change needs no redeploy.
Deno.serve((req) => handle(req, cfg, keySet, db, get));
