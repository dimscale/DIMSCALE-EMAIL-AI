import { assert, assertEquals, assertFalse } from "@std/assert";
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from "jose";
import type { Config } from "../_shared/entra.ts";
import { sha256Hex } from "../_shared/safe.ts";
import { fakeDb } from "../_shared/testing.ts";
import { handle, parsePut } from "./handler.ts";

const T = "11111111-1111-1111-1111-111111111111", C = "22222222-2222-2222-2222-222222222222";
const OID = "33333333-3333-3333-3333-333333333333", OTHER = "99999999-9999-9999-9999-999999999999";
const ORIGIN = "https://localhost:3000";
const cfg: Config = { tenantId: T, clientId: C, allowedOrigins: [ORIGIN] };
const { publicKey, privateKey } = await generateKeyPair("RS256", { extractable: true });
const keySet = createLocalJWKSet({ keys: [{ ...(await exportJWK(publicKey)), kid: "k1", alg: "RS256", use: "sig" }] });

function mint(claims: Record<string, unknown> = {}, aud = C, exp?: number) {
  const now = Math.floor(Date.now() / 1000);
  return new SignJWT({ ver: "2.0", azp: C, scp: "access_as_user", oid: OID, tid: T, ...claims })
    .setProtectedHeader({ alg: "RS256", kid: "k1" }).setIssuer(`https://login.microsoftonline.com/${T}/v2.0`).setAudience(aud)
    .setIssuedAt(now - 100).setExpirationTime(exp ?? now + 3600).sign(privateKey);
}
function setup() {
  const f = fakeDb();
  return { f, call: (token: string | null, init: RequestInit = {}) =>
    handle(new Request("https://f.test/auto-reply-settings", { method: "GET", ...init, headers: { Origin: ORIGIN, ...(token ? { Authorization: `Bearer ${token}` } : {}), ...(init.headers as object) } }), cfg, keySet, f.db) };
}
const put = (body: unknown) => ({ method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
const good = { enabled: true, ownAddress: "Owner@Dimscale.com", contacts: [" Ana@Exemplo.PT ", "rui@cliente.pt"] };
const orig = console.log;
const quiet = (out: string[] = []) => { console.log = (s: string) => out.push(s); return out; };
const loud = () => { console.log = orig; };

Deno.test("AC-36: every negative token is 401 and the database is never touched", async () => {
  const s = setup();
  quiet();
  try {
    const now = Math.floor(Date.now() / 1000);
    const tokens = [
      null, "not-a-jwt",
      await mint({}, "other-aud"),
      await mint({}, C, now - 3600),
      await mint({ scp: "other" }),
      await mint({ ver: "1.0" }),
      await mint({ azp: "someone-else" }),
      await mint({ oid: undefined }),
      await mint({ oid: "o" }),
      await mint({ tid: "44444444-4444-4444-4444-444444444444" }),
    ];
    for (const t of tokens) for (const init of [{}, put(good), { method: "DELETE" }]) {
      const res = await s.call(t, init);
      assertEquals(res.status, 401);
      assertEquals(await res.json(), { ok: false, error: "unauthorized" });
    }
    assertEquals(s.f.calls, []);
    assertEquals((await handle(new Request("https://f.test/x", { headers: { Authorization: "Basic x" } }), cfg, keySet, s.f.db)).status, 401);
  } finally {
    loud();
  }
});

Deno.test("AC-38: OPTIONS needs no token and sets exact CORS; other methods 405; not configured 500", async () => {
  const s = setup();
  quiet();
  try {
    const o = await s.call(null, { method: "OPTIONS" });
    assertEquals(o.status, 204);
    assertEquals(o.headers.get("Access-Control-Allow-Origin"), ORIGIN);
    assertEquals(o.headers.get("Access-Control-Allow-Methods"), "GET, PUT, DELETE, OPTIONS");
    assertEquals(o.headers.get("Access-Control-Allow-Headers"), "authorization, content-type");
    assertEquals(o.headers.get("Vary"), "Origin");
    const evil = await s.call(null, { method: "OPTIONS", headers: { Origin: "https://evil.example" } });
    assertEquals(evil.headers.get("Access-Control-Allow-Origin"), null);
    assertEquals((await s.call(await mint(), { method: "POST" })).status, 405);
    assertEquals((await s.call(await mint(), { method: "PATCH" })).status, 405);
    assertEquals((await handle(new Request("https://f.test/x"), null, null, null)).status, 500);
  } finally {
    loud();
  }
});

Deno.test("AC-37/38: GET, PUT (normalised, atomic, sha of own address), DELETE", async () => {
  const s = setup();
  quiet();
  try {
    const t = await mint();
    let res = await s.call(t);
    assertEquals([res.status, await res.json()], [200, { enabled: false, status: "off", contacts: [] }]);
    res = await s.call(t, put(good));
    assertEquals(res.status, 200);
    assertEquals(res.headers.get("Access-Control-Allow-Origin"), ORIGIN);
    assertEquals(await res.json(), { enabled: true, status: "pending", contacts: ["ana@exemplo.pt", "rui@cliente.pt"] });
    const u = s.f.users.get(OID)!;
    assertEquals(u.sha, await sha256Hex("owner@dimscale.com"));
    // replaced, not merged
    await s.call(t, put({ ...good, contacts: ["x@y.com"] }));
    assertEquals(u.contacts, ["x@y.com"]);
    res = await s.call(t, { method: "DELETE" });
    assertEquals(await res.json(), { enabled: false, status: "off", contacts: [] });
    assertEquals([u.enabled, u.contacts], [false, []]);
  } finally {
    loud();
  }
});

Deno.test("AC-37: invalid PUT bodies are 422 payload_invalid and nothing is written", async () => {
  const s = setup();
  quiet();
  try {
    const t = await mint();
    const many = Array.from({ length: 51 }, (_, i) => `u${i}@x.com`);
    const bad: unknown[] = [
      { ...good, extra: 1 }, { ...good, oid: OTHER }, { enabled: true, contacts: [] }, { ...good, enabled: "yes" }, { ...good, enabled: 1 },
      { ...good, contacts: many }, { ...good, contacts: ["a".repeat(250) + "@x.com"] }, { ...good, contacts: ["nope"] }, { ...good, contacts: ["a b@x.com"] },
      { ...good, contacts: ["a@x"] }, { ...good, contacts: ["a@x.com", " A@X.COM"] }, { ...good, contacts: ["owner@dimscale.com"] },
      { ...good, contacts: [5] }, { ...good, contacts: "a@x.com" }, { ...good, ownAddress: "bad" }, { ...good, ownAddress: 3 }, [], "str", null,
    ];
    for (const b of bad) {
      const res = await s.call(t, put(b));
      assertEquals([res.status, await res.json()], [422, { ok: false, error: "payload_invalid" }], JSON.stringify(b).slice(0, 60));
    }
    assertEquals((await s.call(t, { method: "PUT", body: "{not json" })).status, 422);
    assertEquals(s.f.calls, []);
    assertEquals((await s.call(t, put({ ...good, contacts: many.slice(0, 50) }))).status, 200); // 50 is allowed
  } finally {
    loud();
  }
});

Deno.test("AC-37: rows are keyed by the token oid; a body naming another user has no effect", async () => {
  const s = setup();
  quiet();
  try {
    s.f.addUser(OTHER, ["keep@x.com"], "h", true, "active");
    await s.call(await mint(), put({ ...good, oid: OTHER }));
    assertEquals(s.f.users.get(OTHER)!.contacts, ["keep@x.com"]);
    assertEquals(s.f.users.has(OID), false); // invalid body, nothing written at all
    await s.call(await mint(), put(good));
    assertEquals(s.f.users.get(OTHER)!.contacts, ["keep@x.com"]);
    assertEquals(s.f.users.get(OID)!.contacts.length, 2);
    // GET returns only the caller's data
    const res = await s.call(await mint({ oid: OTHER }));
    assertEquals((await res.json()).contacts, ["keep@x.com"]);
  } finally {
    loud();
  }
});

Deno.test("PUT with enabled but no contacts is stored as off", async () => {
  const s = setup();
  quiet();
  try {
    const res = await s.call(await mint(), put({ enabled: true, ownAddress: "o@d.com", contacts: [] }));
    assertEquals((await res.json()).status, "off");
  } finally {
    loud();
  }
});

Deno.test("database failure is 503 without detail", async () => {
  const s = setup();
  s.f.db.getSettings = () => Promise.reject(new Error("boom with a@b.c"));
  quiet();
  try {
    const res = await s.call(await mint());
    assertEquals([res.status, await res.json()], [503, { ok: false, error: "server_error" }]);
  } finally {
    loud();
  }
});

Deno.test("AC-39: logs hold request id, method, status, reason and count only, never addresses", async () => {
  const s = setup();
  const out = quiet();
  try {
    const t = await mint();
    const a = { enabled: true, ownAddress: "marker-own@secret.example", contacts: ["marker-one@secret.example", "marker-two@secret.example"] };
    await s.call(t, put(a));
    await s.call(t, put({ ...a, contacts: ["marker-bad"] }));
    await s.call(t);
    await s.call(t, { method: "DELETE" });
    await s.call("bad.token.here", put(a));
  } finally {
    loud();
  }
  assert(out.length >= 5);
  for (const l of out) {
    assertFalse(/marker|secret\.example|33333333|Bearer|eyJ/.test(l), l);
    assert(Object.keys(JSON.parse(l)).every((k) => ["fn", "id", "method", "status", "reason", "count", "ms"].includes(k)));
  }
  assertEquals(JSON.parse(out[0]).count, 2);
});

Deno.test("parsePut unit", () => {
  assertEquals(parsePut(good), { enabled: true, ownAddress: "owner@dimscale.com", contacts: ["ana@exemplo.pt", "rui@cliente.pt"] });
  assertEquals(parsePut({ ...good, contacts: [] })?.contacts, []);
});
