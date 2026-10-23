import { errors, jwtVerify, type JWTVerifyGetKey } from "jose";

export type Config = { tenantId: string; clientId: string; allowedOrigins: string[] };

const { JOSEError } = errors;
const SCOPE = "access_as_user";

export const jwksUrl = (tenantId: string) => `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`;

export function readConfig(get: (k: string) => string | undefined): Config | null {
  const tenantId = get("ENTRA_TENANT_ID")?.trim();
  const clientId = get("ENTRA_CLIENT_ID")?.trim();
  const raw = get("ALLOWED_ORIGINS")?.split(",").map((o) => o.trim());
  if (!tenantId || !clientId || !raw?.length || raw.every((o) => !o)) return null;
  const origins: string[] = [];
  for (const o of raw) {
    // Fail closed on anything but a parsable origin (no "*", "null", empty entries).
    if (!o || o === "*" || o === "null") return null;
    try {
      const origin = new URL(o).origin;
      if (origin === "null") return null;
      origins.push(origin);
    } catch {
      return null;
    }
  }
  return { tenantId, clientId, allowedOrigins: origins };
}

// Key-set retrieval failures (network, timeout, non-200, bad JWKS body) are ours, not the caller's.
const isJwksUnavailable = (e: unknown) => {
  const code = (e as { code?: unknown })?.code;
  return !(e instanceof JOSEError) || code === "ERR_JWKS_TIMEOUT" || code === "ERR_JOSE_GENERIC";
};

/** Returns the name on success, else a fixed reason code (never token or claim values). */
export async function validateToken(token: string, keySet: JWTVerifyGetKey, cfg: Config): Promise<{ name: string | null } | { unavailable: true } | string> {
  try {
    const { payload } = await jwtVerify(token, keySet, {
      algorithms: ["RS256"],
      issuer: `https://login.microsoftonline.com/${cfg.tenantId}/v2.0`,
      audience: cfg.clientId,
      clockTolerance: 60,
      requiredClaims: ["exp"],
    });
    if (typeof payload.aud !== "string" || payload.aud !== cfg.clientId) return "aud_mismatch"; // jose accepts arrays
    if (payload.ver !== "2.0") return "ver_mismatch";
    if (payload.azp !== cfg.clientId) return "azp_mismatch";
    if (typeof payload.scp !== "string" || !payload.scp.split(" ").includes(SCOPE)) return "scp_missing";
    return { name: typeof payload.name === "string" ? payload.name : null };
  } catch (e) {
    const code = (e as { code?: unknown }).code;
    if (isJwksUnavailable(e)) return { unavailable: true };
    return typeof code === "string" ? code : "invalid_token";
  }
}
