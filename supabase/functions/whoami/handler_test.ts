import { assertEquals, assertStringIncludes } from "jsr:@std/assert@1";
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from "jose";
import { type Config, handle, readConfig } from "./handler.ts";

const T = "11111111-1111-1111-1111-111111111111";
const C = "22222222-2222-2222-2222-222222222222";
const ORIGIN = "https://localhost:3000";
const cfg: Config = { tenantId: T, clientId: C, allowedOrigins: [ORIGIN] };

const { publicKey, privateKey } = await generateKeyPair("RS256", { extractable: true });
const jwk = { ...(await exportJWK(publicKey)), kid: "k1", alg: "RS256", use: "sig" };
const keySet = createLocalJWKSet({ keys: [jwk] });

type Over = { claims?: Record<string, unknown>; iss?: string; aud?: string; exp?: number; nbf?: number; kid?: string; key?: CryptoKey };
async function mint(o: Over = {}): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const claims: Record<string, unknown> = { ver: "2.0", azp: C, scp: "access_as_user", name: "Ada Test", oid: "o", tid: T, preferred_username: "a@b.c", ...o.claims };
  return await new SignJWT(claims)
    .setProtectedHeader({ alg: "RS256", kid: o.kid ?? "k1" })
    .setIssuer(o.iss ?? `https://login.microsoftonline.com/${T}/v2.0`)
    .setAudience(o.aud ?? C)
    .setIssuedAt(now - 100)
    .setNotBefore(o.nbf ?? now - 100)
    .setExpirationTime(o.exp ?? now + 3600)
    .sign(o.key ?? privateKey);
}
const call = (token: string | null, init: RequestInit = {}, c: Config | null = cfg) =>
  handle(new Request("https://f.test/whoami", { method: "GET", ...init, headers: { Origin: ORIGIN, ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(init.headers as object) } }), c, c ? keySet : null);

async function assert401(res: Response) {
  assertEquals(res.status, 401);
  assertEquals(await res.json(), { ok: false, error: "unauthorized" });
}
const b64 = (o: object) => btoa(JSON.stringify(o)).replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");

Deno.test("AC-8 missing / non-bearer / malformed", async () => {
  await assert401(await call(null));
  await assert401(await call(null, { headers: { Authorization: "Basic abc" } }));
  await assert401(await call(null, { headers: { Authorization: "Bearer" } }));
  await assert401(await call("not-a-jwt"));
});

Deno.test("AC-9 bad signature, alg none/HS256, unknown kid", async () => {
  const other = await generateKeyPair("RS256");
  await assert401(await call(await mint({ key: other.privateKey })));
  const none = `${b64({ alg: "none" })}.${b64({ iss: `https://login.microsoftonline.com/${T}/v2.0`, aud: C })}.`;
  await assert401(await call(none));
  const hs = await new SignJWT({ ver: "2.0" }).setProtectedHeader({ alg: "HS256", kid: "k1" }).setIssuer(`https://login.microsoftonline.com/${T}/v2.0`).setAudience(C).setExpirationTime("1h").sign(new TextEncoder().encode("x".repeat(32)));
  await assert401(await call(hs));
  await assert401(await call(await mint({ kid: "unknown" })));
});

Deno.test("AC-10 issuer and ver", async () => {
  for (const iss of [`https://login.microsoftonline.com/${C}/v2.0`, `https://sts.windows.net/${T}/`, `https://login.microsoftonline.com/${T}/v2.0/`, `https://login.microsoftonline.com/${T}`]) {
    await assert401(await call(await mint({ iss })));
  }
  await assert401(await call(await mint({ claims: { ver: "1.0" } })));
});

Deno.test("AC-11 aud, azp, scp", async () => {
  for (const aud of [`api://${C}`, "00000003-0000-0000-c000-000000000000", "33333333-3333-3333-3333-333333333333"]) {
    await assert401(await call(await mint({ aud })));
  }
  await assert401(await call(await mint({ claims: { azp: "33333333-3333-3333-3333-333333333333" } })));
  await assert401(await call(await mint({ claims: { scp: "other_scope" } })));
  await assert401(await call(await mint({ claims: { scp: undefined } })));
});

Deno.test("AC-12 expiry and not-before with 60 s tolerance", async () => {
  const now = Math.floor(Date.now() / 1000);
  await assert401(await call(await mint({ exp: now - 120 })));
  assertEquals((await call(await mint({ exp: now - 30 }))).status, 200);
  await assert401(await call(await mint({ nbf: now + 120 })));
  assertEquals((await call(await mint({ nbf: now + 30 }))).status, 200);
});

Deno.test("AC-13 success body has only ok and name", async () => {
  const res = await call(await mint());
  assertEquals(res.status, 200);
  assertEquals(await res.json(), { ok: true, name: "Ada Test" });
  const noName = await call(await mint({ claims: { name: undefined } }));
  assertEquals(await noName.json(), { ok: true, name: null });
});

Deno.test("AC-14 CORS", async () => {
  const pre = await call(null, { method: "OPTIONS" });
  assertEquals(pre.status, 204);
  assertEquals(pre.headers.get("Access-Control-Allow-Origin"), ORIGIN);
  assertEquals(pre.headers.get("Access-Control-Allow-Headers"), "authorization");
  assertEquals(pre.headers.get("Access-Control-Allow-Methods"), "GET, OPTIONS");
  assertEquals(pre.headers.get("Vary"), "Origin");
  const bad = await call(null, { method: "OPTIONS", headers: { Origin: "https://evil.test" } });
  assertEquals(bad.headers.get("Access-Control-Allow-Origin"), null);
  const ok = await call(await mint());
  assertEquals(ok.headers.get("Access-Control-Allow-Origin"), ORIGIN);
  const un = await call(null);
  assertEquals(un.headers.get("Access-Control-Allow-Origin"), ORIGIN);
  const evil200 = await call(await mint(), { headers: { Origin: "https://evil.test" } });
  assertEquals(evil200.headers.get("Access-Control-Allow-Origin"), null);
  for (const method of ["POST", "PUT", "DELETE"]) {
    assertEquals((await call(await mint(), { method })).status, 405);
  }
});

Deno.test("fails closed without config, generic body", async () => {
  const res = await call(await mint(), {}, null);
  assertEquals(res.status, 500);
  assertEquals(await res.json(), { ok: false, error: "server_error" });
  assertEquals(readConfig(() => undefined), null);
  assertEquals(readConfig((k) => ({ ENTRA_TENANT_ID: T, ENTRA_CLIENT_ID: C, ALLOWED_ORIGINS: " " })[k]), null);
  assertStringIncludes(JSON.stringify(readConfig((k) => ({ ENTRA_TENANT_ID: T, ENTRA_CLIENT_ID: C, ALLOWED_ORIGINS: `${ORIGIN}, https://b.test` })[k])), "https://b.test");
});

Deno.test("token without exp is rejected", async () => {
  const now = Math.floor(Date.now() / 1000);
  const t = await new SignJWT({ ver: "2.0", azp: C, scp: "access_as_user" })
    .setProtectedHeader({ alg: "RS256", kid: "k1" })
    .setIssuer(`https://login.microsoftonline.com/${T}/v2.0`)
    .setAudience(C)
    .setIssuedAt(now - 100)
    .sign(privateKey);
  await assert401(await call(t));
});

Deno.test("ALLOWED_ORIGINS is normalized and fails closed on wildcard/null/garbage", () => {
  const rc = (o: string) => readConfig((k) => ({ ENTRA_TENANT_ID: T, ENTRA_CLIENT_ID: C, ALLOWED_ORIGINS: o })[k]);
  assertEquals(rc("https://Localhost:3000/, HTTPS://B.test")?.allowedOrigins, ["https://localhost:3000", "https://b.test"]);
  for (const bad of ["*", "null", "https://a.test,*", "https://a.test,,https://b.test", "not a url", "data:text/plain,x"]) assertEquals(rc(bad), null);
});
