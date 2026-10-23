import { importPKCS8, SignJWT } from "jose";

export type AppTokenCfg = { tenantId: string; clientId: string; x5t: string; privateKeyB64: string };

/** Client credentials with a certificate assertion (PS256). The token is cached in memory until 5 min before expiry. */
export function makeAppTokenProvider(cfg: AppTokenCfg, fetchFn: typeof fetch = fetch, nowMs: () => number = Date.now) {
  let cached: { token: string; refreshAt: number } | null = null;
  let key: CryptoKey | null = null;
  const endpoint = `https://login.microsoftonline.com/${cfg.tenantId}/oauth2/v2.0/token`;

  async function assertion(): Promise<string> {
    // The secret is base64 of the PKCS#8 DER on one line; wrap it as PEM for jose.
    key ??= await importPKCS8(`-----BEGIN PRIVATE KEY-----\n${cfg.privateKeyB64}\n-----END PRIVATE KEY-----`, "PS256");
    const now = Math.floor(nowMs() / 1000);
    return await new SignJWT({})
      .setProtectedHeader({ alg: "PS256", typ: "JWT", "x5t#S256": cfg.x5t })
      .setAudience(endpoint)
      .setIssuer(cfg.clientId)
      .setSubject(cfg.clientId)
      .setJti(crypto.randomUUID())
      .setNotBefore(now)
      .setIssuedAt(now)
      .setExpirationTime(now + 300)
      .sign(key);
  }

  /** Throws Error("token_unavailable") on any failure (no detail: it could carry secrets). */
  async function getToken(): Promise<string> {
    if (cached && nowMs() < cached.refreshAt) return cached.token;
    try {
      const res = await fetchFn(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: cfg.clientId,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
          client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
          client_assertion: await assertion(),
        }),
        signal: AbortSignal.timeout(10_000),
      });
      if (!res.ok) throw new Error("status");
      const j = await res.json() as { access_token?: unknown; expires_in?: unknown };
      if (typeof j.access_token !== "string" || typeof j.expires_in !== "number") throw new Error("shape");
      cached = { token: j.access_token, refreshAt: nowMs() + Math.max(0, j.expires_in - 300) * 1000 };
      return cached.token;
    } catch {
      throw new Error("token_unavailable");
    }
  }
  return { getToken, invalidate: () => { cached = null; } };
}

export function readAppTokenCfg(get: (k: string) => string | undefined): AppTokenCfg | null {
  const tenantId = get("ENTRA_TENANT_ID")?.trim(), clientId = get("GRAPH_APP_CLIENT_ID")?.trim();
  const x5t = get("GRAPH_CERT_X5T_S256")?.trim(), privateKeyB64 = get("GRAPH_CERT_PRIVATE_KEY")?.trim();
  return tenantId && clientId && x5t && privateKeyB64 ? { tenantId, clientId, x5t, privateKeyB64 } : null;
}
