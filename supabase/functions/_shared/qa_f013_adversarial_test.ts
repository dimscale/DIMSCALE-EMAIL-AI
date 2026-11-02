// QA (adversarial) for F-013 amendment (AC-44 to AC-47, AC-49). Mocked Graph only, no network, no DB.
// Tests named "KNOWN LIMIT" pin current behaviour so a change is noticed.
import { assert, assertEquals, assertFalse } from "@std/assert";
import { readSettings } from "./config.ts";
import { authFailed, evaluate, type FilterCtx, isNoReplyLocalPart } from "./filters.ts";
import { makeGraph, type Msg } from "./graph.ts";
import { sha256Hex } from "./safe.ts";

const OWN = "owner@dimscale.com";
const NOW = Date.parse("2026-10-23T12:00:00Z");
const ownSha = await sha256Hex(OWN);
const S = readSettings(() => undefined).settings;
const ctx = (over: Partial<FilterCtx> = {}): FilterCtx => ({ mailSha256: ownSha, inboxFolderId: "INBOX", nowMs: NOW, countsToday: { user: 0, global: 0 }, ...over });
const msg = (over: Partial<Msg> = {}): Msg => ({
  odataType: "#microsoft.graph.message", subject: "Hi", fromAddress: "ana@exemplo.pt", fromName: "Ana", toAddresses: [OWN], isDraft: false,
  inference: "focused", receivedDateTime: "2026-10-23T11:00:00Z", headers: {}, body: "Hello", parentFolderId: "INBOX", ...over,
});
const withAr = (value: string | undefined, over: Partial<Msg> = {}) =>
  msg({ headers: value === undefined ? {} : { "authentication-results": value }, ...over });

Deno.test("ADV F-013/AC-45: dmarc=fail and compauth=fail token boundaries (RFC-style values)", () => {
  const skip = [
    "dmarc=fail", "DMARC=FAIL", "Compauth=FAIL reason=000", "spf=pass; dmarc=fail action=none header.from=x.pt",
    "spf=pass;dmarc=fail", "dmarc=pass\tdmarc=fail", "dmarc=pass\ndmarc=fail", "dmarc=fail;", "dmarc=fail(p=none)",
    "mx.example;\r\n\tspf=pass;\r\n\tdmarc=fail header.from=x.pt", "dmarc=pass; compauth=fail reason=000",
    "dmarc=fail (p=reject)", "dmarc=fail (a (nested) b)",
    "dmarc=pass reason=x; dmarc=fail",
  ];
  for (const v of skip) assert(authFailed(v), `should skip: ${v}`);
  const pass = [
    "dmarc=pass; spf=fail", "dkim=fail", "spf=fail", "dmarc=pass reason=\"dmarc=fail\"", "dmarc=pass (policy \"dmarc=fail\")",
    "dmarc=pass (note: dmarc=fail)", "dmarc=pass (note dmarc=fail here)", "dmarc=pass (a (dmarc=fail) b)", "dmarc=pass (unclosed dmarc=fail", "dmarc=fail)", "dmarc=fail.", "xcompauth=fail", "compauth=fail_x", "dmarc=fail_x",
    "dmarc=failx", "xdmarc=fail", "dmarc = fail", "dmarc=\r\n fail", "dmarc:fail", "",
  ];
  for (const v of pass) assertFalse(authFailed(v), `should pass: ${JSON.stringify(v)}`);
  assertFalse(authFailed(undefined));
});

Deno.test("ADV F-013/AC-45/46: the Graph mapper keeps every Authentication-Results header; name case and folding handled", async () => {
  const headersOf = async (internetMessageHeaders: unknown) => {
    const fetchFn = (() => Promise.resolve(new Response(JSON.stringify({ isDraft: false, internetMessageHeaders }), { status: 200 }))) as unknown as typeof fetch;
    const g = makeGraph({ getToken: () => Promise.resolve("TOK"), invalidateToken: () => {}, fetchFn, webhookUrl: "https://x.test/hook", clientState: "STATE", nowMs: () => NOW });
    const m = await g.readMessage("33333333-3333-3333-3333-333333333333", "id");
    assert(m.ok);
    return authFailed(m.data.headers["authentication-results"]);
  };
  assert(await headersOf([{ name: "AUTHENTICATION-RESULTS", value: "mx; DMARC=FAIL" }]), "upper-case name");
  assert(await headersOf([{ name: "Authentication-Results", value: "dmarc=pass" }, { name: "authentication-results", value: "compauth=fail reason=000" }]), "second header fails");
  assert(await headersOf([{ name: "Authentication-Results", value: "dmarc=fail" }, { name: "Authentication-Results", value: "dmarc=pass" }]), "first header fails");
  assertFalse(await headersOf([{ name: "Authentication-Results", value: "spf=fail" }, { name: "Authentication-Results", value: "dmarc=pass" }]), "spf only, second pass");
  assert(await headersOf([{ name: "Authentication-Results", value: "mx.example;\r\n\tspf=pass;\r\n\tdmarc=fail(p=none)" }]), "folded value");
  assertFalse(await headersOf([{ name: "X-Authentication-Results", value: "dmarc=fail" }]), "different header name");
  assertFalse(await headersOf([{ name: "Authentication-Results", value: null }]), "null value fails open");
  assertFalse(await headersOf([{ name: 7, value: "dmarc=fail" }, "junk", null]), "malformed entries ignored");
  assertFalse(await headersOf(undefined), "no headers at all");
  assertFalse(await headersOf([]), "empty headers");
});

Deno.test("ADV F-013/AC-44/45: filter order around auth_failed (each filter keeps its own code when it is the only one)", async () => {
  const dmarc = "dmarc=fail";
  const cases: [string, Msg, Partial<FilterCtx>, typeof S][] = [
    ["auto_submitted", withAr(dmarc, { headers: { "authentication-results": dmarc, "auto-submitted": "auto-replied" } }), {}, S],
    ["noreply_sender", withAr(dmarc, { fromAddress: "noreply@exemplo.pt" }), {}, S],
    ["focused_other", withAr(dmarc, { inference: "other" }), {}, { ...S, filterFocusedOther: true }],
    ["auth_failed", withAr(dmarc, { inference: "other" }), {}, S],
    ["auth_failed", withAr(dmarc, { receivedDateTime: "2026-10-22T00:00:00Z" }), {}, S], // before too_old (13)
    ["auth_failed", withAr(dmarc, { body: "   " }), {}, S], // before empty_body (14)
    ["auth_failed", withAr(dmarc, { fromAddress: "x@spam.example" }), {}, { ...S, skipSenderDomains: ["spam.example"] }], // before 15
    ["auth_failed", withAr(dmarc), { countsToday: { user: 99, global: 0 } }, S], // before cap (17)
    ["from_self", withAr(dmarc, { fromAddress: OWN }), {}, S],
    ["not_inbox_message", withAr(dmarc, { parentFolderId: "SENT" }), {}, S],
    ["not_direct_recipient", withAr(dmarc, { toAddresses: [] }), {}, S],
    ["mailing_list", withAr(dmarc, { headers: { "authentication-results": dmarc, "list-id": "<l.x>" } }), {}, S], // filter 10 before 18
    ["too_old", withAr(undefined, { receivedDateTime: "2026-10-22T00:00:00Z" }), {}, S], // no header, filter 18 passes
    ["auth_failed", withAr(dmarc, { receivedDateTime: "2026-10-22T00:00:00Z" }), {}, { ...S, filterAuthFailed: true }],
  ];
  for (const [code, m, over, s] of cases) {
    assertEquals(await evaluate(m, ctx(over), s), { skip: code }, `expected ${code}`);
  }
  // Toggle off: the same too-old message gets its own code, auth is not consulted.
  assertEquals(await evaluate(withAr(dmarc, { receivedDateTime: "2026-10-22T00:00:00Z" }), ctx(), { ...S, filterAuthFailed: false }), { skip: "too_old" });
  // Toggle off with focused other on: the Focused Other code is still given.
  assertEquals(await evaluate(withAr(dmarc, { inference: "other" }), ctx(), { ...S, filterAuthFailed: false, filterFocusedOther: true }), { skip: "focused_other" });
});

Deno.test("ADV F-013/AC-47/AC-13: FILTER_AUTH_FAILED and FILTER_FOCUSED_OTHER parse only exact lower-case true/false", () => {
  // config.ts bool(): "TRUE"/"FALSE" are invalid, so the default applies and config_default is logged.
  const r = (k: string, v: string) => readSettings((x) => (x === k ? v : undefined));
  assert(r("FILTER_AUTH_FAILED", "FALSE").settings.filterAuthFailed, "FALSE is invalid: default (on)");
  assertFalse(r("FILTER_AUTH_FAILED", "false").settings.filterAuthFailed);
  assertFalse(r("FILTER_FOCUSED_OTHER", "TRUE").settings.filterFocusedOther, "TRUE is invalid: default (off)");
  assert(r("FILTER_FOCUSED_OTHER", "true").settings.filterFocusedOther);
  assert(r("FILTER_AUTH_FAILED", "FALSE").defaulted.includes("FILTER_AUTH_FAILED"));
});

Deno.test("ADV F-013/AC-49: no-reply local parts (filter 11) are matched exactly and case-insensitively", () => {
  for (const a of ["noreply@x.pt", "NOREPLY@x.pt", "No-Reply@x.pt", "donotreply@x.pt", "DO-NOT-REPLY@x.pt", "mailer-daemon@x.pt", "Postmaster@x.pt"]) {
    assert(isNoReplyLocalPart(a), a);
  }
  for (const a of ["ana@x.pt", "noreply.team@x.pt", "x-noreply@x.pt", "noreplyx@x.pt"]) assertFalse(isNoReplyLocalPart(a), a);
  // Amendment 2026-11-03: NFKC, ignorable characters, lower-case, trim and the "+" cut are normalised.
  for (const a of ["noreply+x@x.pt", "ｎｏｒｅｐｌｙ@x.pt" /* fullwidth */, "noreply​@x.pt" /* zero-width */, "no­reply@x.pt", " noreply@x.pt", "no-reply+a+b@x.pt", "ｎｏｒｅｐｌｙ＋x@x.pt"]) {
    assert(isNoReplyLocalPart(a), `no-reply: ${JSON.stringify(a)}`);
  }
  // KNOWN LIMIT: look-alikes from other alphabets are not handled. Pinned.
  assertFalse(isNoReplyLocalPart("nоreply@x.pt" /* Cyrillic o */), "pinned: Cyrillic look-alike");
});

Deno.test("ADV F-013/AC-47: header values reach no log line, DB write or OpenAI input when auth_failed triggers (evaluate only)", async () => {
  // evaluate() reads headers and returns a code only; this asserts the verdict carries no header text.
  const MARK = "AR-MARK-7c1";
  const v = await evaluate(withAr(`mx.${MARK}; dmarc=fail header.from=${MARK}.pt`), ctx(), S);
  assertEquals(v, { skip: "auth_failed" });
  assertFalse(JSON.stringify(v).includes(MARK));
});

Deno.test("ADV F-013 (performance): a 1 MB Authentication-Results value is matched in linear time", () => {
  const big = "dmarc=pass ".repeat(90_000) + "x";
  const t0 = performance.now();
  const r = authFailed(big);
  const ms = performance.now() - t0;
  assertFalse(r);
  assert(ms < 1000, `took ${ms.toFixed(0)} ms`);
});
