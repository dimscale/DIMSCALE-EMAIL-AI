import { createRemoteJWKSet } from "jose";
import { handle, jwksUrl, readConfig } from "./handler.ts";

const cfg = readConfig((k) => Deno.env.get(k));
// Created once: jose caches keys and refetches on an unknown kid (key rotation).
const keySet = cfg ? createRemoteJWKSet(new URL(jwksUrl(cfg.tenantId))) : null;

Deno.serve((req) => handle(req, cfg, keySet));
