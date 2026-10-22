// QA additions for F-002 (independent check). Local key pair and JWKS only, no network.
import { assert, assertEquals, assertFalse } from "jsr:@std/assert@1";
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from "jose";
import { type Config, handle, readConfig } from "./handler.ts";

const T = "11111111-1111-1111-1111-111111111111";
const C = "22222222-2222-2222-2222-222222222222";
const ORIGIN = "https://localhost:3000";
const ISS = `https://login.microsoftonline.com/${T}/v2.0`;
const cfg: Config = { tenantId: T, clientId: C, allowedOrigins: [ORIGIN] };
const UNAUTHORIZED = '{"ok":false,"error":"unauthorized"}';

const { publicKey, privateKey } = await generateKeyPair("RS256", { extractable: true });
const keySet = createLocalJWKSet({ keys: [{ ...(await exportJWK(publicKey)), kid: "k1", alg: "RS256", use: "sig" }] });

type Opts = { aud?: string | string[]; iss?: string; exp?: number; nbf?: number };
async function mint(claims: Record<string, unknown> = {}, o: Opts = {}): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const payload = { ver: "2.0", azp: C, scp: "access_as_user", name: "Ada Test", oid: "oid-123", tid: T, preferred_username: "ada@example.com", ...claims };
  return await new SignJWT(payload)
    .setProtectedHeader({ alg: "RS256", kid: "k1" })
    .setIssuer(o.iss ?? ISS)
    .setAudience(o.aud ?? C)
    .setIssuedAt(now - 100)
    .setNotBefore(o.nbf ?? now - 100)
    .setExpirationTime(o.exp ?? now + 3600)
    .sign(privateKey);
}
const req = (headers: Record<string, string> = {}, init: RequestInit = {}) =>
  new Request("https://f.test/whoami", { method: "GET", ...init, headers: { Origin: ORIGIN, ...headers } });
const withBearer = (t: string) => ({ Authorization: `Bearer ${t}` });

Deno.test("AC-11 scp: access_as_user accepted when other scopes are present", async () => {
  for (const scp of ["openid profile access_as_user", "access_as_user User.Read", "openid access_as_user offline_access"]) {
    const res = await handle(req(withBearer(await mint({ scp }))), cfg, keySet);
    assertEquals(res.status, 200, scp);
  }
});

Deno.test("AC-11 scp: look-alike and empty scopes rejected", async () => {
  for (const scp of ["access_as_user_x", "access_as_user_x openid", "xaccess_as_user", "Access_As_User", "access_as", ""]) {
    const res = await handle(req(withBearer(await mint({ scp }))), cfg, keySet);
    assertEquals(res.status, 401, scp);
  }
  // scp must be the space-separated string Entra emits, not an array
  assertEquals((await handle(req(withBearer(await mint({ scp: ["access_as_user"] }))), cfg, keySet)).status, 401);
});

Deno.test("AC-10/11 missing ver, azp or scp claim rejected", async () => {
  for (const claims of [{ ver: undefined }, { azp: undefined }, { scp: undefined }]) {
    const res = await handle(req(withBearer(await mint(claims))), cfg, keySet);
    assertEquals(res.status, 401, JSON.stringify(Object.keys(claims)));
    assertEquals(await res.text(), UNAUTHORIZED);
  }
});

Deno.test("AC-8 Authorization must be exactly 'Bearer <token>'", async () => {
  const t = await mint();
  for (const auth of [`bearer ${t}`, `BEARER ${t}`, `Bearer  ${t}`, `Bearer ${t} extra`, `Token ${t}`, `Bearer ${t},${t}`]) {
    const res = await handle(req({ Authorization: auth }), cfg, keySet);
    assertEquals(res.status, 401, auth.slice(0, 12));
    assertEquals(await res.text(), UNAUTHORIZED);
  }
});

Deno.test("AC-8/10/11 a 401 body is the fixed JSON and never echoes the token, reason or claims", async () => {
  const cases: [string, string | undefined][] = [
    ["no header", undefined],
    ["garbage token", "Bearer abc.def.ghi"],
    ["wrong issuer", await mint({}, { iss: `https://sts.windows.net/${T}/` })],
    ["wrong audience", await mint({}, { aud: `api://${C}` })],
    ["wrong scope", await mint({ scp: "other_scope" })],
    ["expired", await mint({}, { exp: Math.floor(Date.now() / 1000) - 300 })],
  ];
  for (const [label, value] of cases) {
    const headers = value === undefined ? {} : value.startsWith("Bearer ") ? { Authorization: value } : withBearer(value);
    const res = await handle(req(headers), cfg, keySet);
    const text = await res.text();
    assertEquals(res.status, 401, label);
    assertEquals(text, UNAUTHORIZED, label);
    if (value) for (const part of value.replace("Bearer ", "").split(".")) assertFalse(text.includes(part), label);
    for (const leak of ["ERR_", "iss", "aud", "scp", "ver", "expired", "claim", "Ada", "ada@"]) assertFalse(text.includes(leak), `${label}: ${leak}`);
  }
});

Deno.test("AC-14 without ALLOWED_ORIGINS (null config): 500, no CORS headers, no ACAO on any method", async () => {
  const res = await handle(req(withBearer(await mint())), null, null);
  assertEquals(res.status, 500);
  assertEquals(res.headers.get("Access-Control-Allow-Origin"), null);
  assertEquals(await res.text(), '{"ok":false,"error":"server_error"}');
  const pre = await handle(req({}, { method: "OPTIONS" }), null, null);
  assertEquals(pre.status, 500);
  assertEquals(pre.headers.get("Access-Control-Allow-Origin"), null);
  assertEquals(pre.headers.get("Access-Control-Allow-Headers"), null);
});

Deno.test("config: each required env value missing fails closed", () => {
  const env: Record<string, string> = { ENTRA_TENANT_ID: T, ENTRA_CLIENT_ID: C, ALLOWED_ORIGINS: ORIGIN };
  assert(readConfig((k) => env[k]));
  for (const k of Object.keys(env)) {
    assertEquals(readConfig((x) => (x === k ? undefined : env[x])), null, k);
    assertEquals(readConfig((x) => (x === k ? "  " : env[x])), null, `${k} blank`);
  }
});

Deno.test("AC-14 a wildcard in ALLOWED_ORIGINS never produces Access-Control-Allow-Origin", async () => {
  const wild: Config = { ...cfg, allowedOrigins: ["*"] };
  const res = await handle(req({ ...withBearer(await mint()), Origin: "https://evil.test" }), wild, keySet);
  assertEquals(res.status, 200);
  assertEquals(res.headers.get("Access-Control-Allow-Origin"), null);
});

Deno.test("AC-14 a request without Origin (non-browser) is answered without CORS headers", async () => {
  const res = await handle(new Request("https://f.test/whoami", { headers: withBearer(await mint()) }), cfg, keySet);
  assertEquals(res.status, 200);
  assertEquals(res.headers.get("Access-Control-Allow-Origin"), null);
});

Deno.test("AC-14 methods other than GET/OPTIONS return 405 even with a valid token", async () => {
  for (const method of ["PATCH", "DELETE", "PUT", "POST"]) {
    const res = await handle(req(withBearer(await mint()), { method }), cfg, keySet);
    assertEquals(res.status, 405, method);
  }
});

Deno.test("AC-13/14 no response body contains the token, the Authorization value or non-name claims", async () => {
  const t = await mint();
  const ok = await handle(req(withBearer(t)), cfg, keySet);
  const okText = await ok.text();
  assertEquals(okText, '{"ok":true,"name":"Ada Test"}');
  for (const part of t.split(".")) assertFalse(okText.includes(part));
  for (const leak of ["oid-123", T, "ada@example.com", "preferred_username", "tid", "oid"]) assertFalse(okText.includes(leak), leak);
  const pre = await handle(req({}, { method: "OPTIONS" }), cfg, keySet);
  assertEquals(await pre.text(), "");
});

Deno.test("AC-16 logs contain only id, status, reason and ms, never token or claim values", async () => {
  const t = await mint();
  const lines: string[] = [];
  const orig = console.log;
  console.log = (...args: unknown[]) => {
    lines.push(args.map(String).join(" "));
  };
  try {
    await handle(req(withBearer(t)), cfg, keySet);
    await handle(req(withBearer(await mint({ scp: "x" }))), cfg, keySet);
    await handle(req(withBearer(await mint({}, { iss: `${ISS}/` }))), cfg, keySet);
    await handle(req({}, { method: "OPTIONS" }), cfg, keySet);
    await handle(req({}, { method: "POST" }), cfg, keySet);
    await handle(req(withBearer(t)), null, null);
  } finally {
    console.log = orig;
  }
  assertEquals(lines.length, 6);
  const secrets = [...t.split("."), "Ada Test", "oid-123", "ada@example.com", T, C, "Bearer"];
  for (const line of lines) {
    assertEquals(Object.keys(JSON.parse(line)).sort(), ["id", "ms", "reason", "status"]);
    for (const s of secrets) assertFalse(line.includes(s), s);
  }
});

Deno.test("AC-9 a JWKS fetch outage is reported as a server error, not as a rejected sign-in", async () => {
  // Known-risk probe: today jose's fetch failure has no code, so handler maps it to 401.
  const down = (async () => {
    throw new Error("network down");
  }) as never;
  const res = await handle(req(withBearer(await mint())), cfg, down);
  assert(res.status >= 500, `expected 5xx, got ${res.status}`);
  assertFalse((await res.text()).includes("network down"));
});

Deno.test("AC-11 an aud array that merely contains the client ID is rejected (spec: exact match)", async () => {
  const res = await handle(req(withBearer(await mint({}, { aud: [C, "33333333-3333-3333-3333-333333333333"] }))), cfg, keySet);
  assertEquals(res.status, 401);
});
