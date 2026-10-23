import { decodeJwt, type JWTVerifyGetKey } from "jose";
import { type Config, validateToken } from "./entra.ts";

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** F-002 validation plus the extra checks for functions that store per-user data (F-013/AC-36).
 *  Returns the lower-cased oid, a fixed reason code, or { unavailable }. */
export async function validateStateful(token: string, keySet: JWTVerifyGetKey, cfg: Config): Promise<{ oid: string } | { unavailable: true } | string> {
  const r = await validateToken(token, keySet, cfg);
  if (typeof r === "string" || "unavailable" in r) return r;
  const p = decodeJwt(token); // signature already verified above
  if (p.tid !== cfg.tenantId) return "tid_mismatch";
  if (typeof p.oid !== "string" || !GUID.test(p.oid)) return "oid_invalid";
  return { oid: p.oid.toLowerCase() };
}
