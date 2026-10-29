import { assert, assertEquals, assertFalse } from "@std/assert";
import { buildInput } from "../_shared/ai.ts";
import { generateSummary, hasLink, parseSummary, SUMMARY_INSTRUCTIONS, SUMMARY_SCHEMA } from "./summary.ts";

const DI = { subject: "Fatura", fromName: "Ana", fromDomain: "exemplo.pt", body: "Pague em https://exemplo.pt/pay ou ana@exemplo.pt" };
const INPUT = buildInput(DI);
const good = { points: ["a", "b", "c"], type: "request", priority: "high", tone: "neutral", deadline: "sexta", requested_action: null };
const wrap = (t: string) => ({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: t }] }] });
const out = (o: unknown) => wrap(JSON.stringify(o));
const invalid = { ok: false, reason: "ai_output_invalid", retryable: false } as const;
const run = (o: unknown) => parseSummary(o, INPUT);

Deno.test("AC-9: valid summary maps requested_action to requestedAction", () => {
  assertEquals(run(out(good)), { ok: true, summary: { points: ["a", "b", "c"], type: "request", priority: "high", tone: "neutral", deadline: "sexta", requestedAction: null } });
  assert(run(out({ ...good, deadline: null, requested_action: "Pagar" })).ok);
});

Deno.test("AC-9: a URL or address is rejected even when the email had it, and a summary without links passes", () => {
  assert(run(out({ ...good, points: ["a", "b", "c"] })).ok);
  for (const o of [
    { ...good, points: ["Ver https://exemplo.pt/pay", "b", "c"] },
    { ...good, points: ["a", "ana@exemplo.pt", "c"] },
    { ...good, points: ["a", "b", "www.exemplo.pt"] },
    { ...good, deadline: "http://exemplo.pt" },
    { ...good, requested_action: "ver //exemplo.pt/pay" },
  ]) assertEquals(run(out(o)), invalid);
});

Deno.test("hasLink: URL, www., protocol-relative and address tokens", () => {
  for (const s of ["see http://evil.example", "HTTPS://X.y", "www.evil.example", "go //evil.example/x", "boss@evil.example"]) assert(hasLink(s), s);
  for (const s of ["Pagar a fatura até sexta.", "a/b e c/d", "Olá @Ana", "2026-10-29"]) assertFalse(hasLink(s), s);
});

Deno.test("AC-9: every failure is ai_output_invalid", () => {
  const msg = (content: unknown[]) => ({ status: "completed", output: [{ type: "message", content }] });
  const ot = { type: "output_text", text: JSON.stringify(good) };
  const cases: unknown[] = [
    { ...out(good), status: "incomplete" },
    msg([{ type: "refusal", refusal: "no" }]),
    msg([ot, ot]),
    msg([]),
    wrap("not json"), wrap("[]"), wrap("null"),
    out({ ...good, extra: 1 }),
    out({ ...good, points: undefined }),
    out({ ...good, points: ["a", "b"] }),
    out({ ...good, points: ["a", "b", "c", "d"] }),
    out({ ...good, points: ["a", " ", "c"] }),
    out({ ...good, points: ["a", "b", "x".repeat(301)] }),
    out({ ...good, points: ["a", "b", 3] }),
    out({ ...good, type: "phishing" }),
    out({ ...good, priority: "urgent" }),
    out({ ...good, tone: "angry" }),
    out({ ...good, deadline: "" }),
    out({ ...good, deadline: "x".repeat(201) }),
    out({ ...good, requested_action: "" }),
    out({ ...good, requested_action: 5 }),
    out({ ...good, points: ["<b>x</b>", "b", "c"] }),
    out({ ...good, points: ["a > b", "b", "c"] }),
    out({ ...good, deadline: "a<b" }),
    out({ ...good, requested_action: "a>b" }),
    out({ ...good, points: ["Veja https://evil.example/x", "b", "c"] }),
    out({ ...good, points: ["a", "escreva para boss@evil.example", "c"] }),
    out({ ...good, requested_action: "ver //evil.example/x" }),
    null, "x",
  ];
  for (const c of cases) assertEquals(run(c), invalid);
});

Deno.test("AC-7: generateSummary sends the exact request", async () => {
  let seen: { url: string; init: RequestInit } | undefined;
  const f = ((url: string, init: RequestInit) => (seen = { url, init }, Promise.resolve(new Response(JSON.stringify(out(good)))))) as unknown as typeof fetch;
  const cfg = { apiKey: "KEY", model: "gpt-6-luna", maxOutputTokens: 2000, reasoningEffort: "low" };
  const r = await generateSummary({ ...DI, body: "Pague em https://exemplo.pt/pay" }, cfg, f);
  assert(r.ok);
  assertEquals(seen!.url, "https://api.openai.com/v1/responses");
  assertEquals((seen!.init.headers as Record<string, string>).Authorization, "Bearer KEY");
  assert(seen!.init.signal);
  const b = JSON.parse(seen!.init.body as string);
  assertEquals(Object.keys(b).sort(), ["input", "instructions", "max_output_tokens", "model", "reasoning", "store", "text"]);
  assertFalse("tools" in b);
  assertEquals([b.model, b.store, b.max_output_tokens, b.reasoning], ["gpt-6-luna", false, 2000, { effort: "low" }]);
  assertEquals(b.instructions, SUMMARY_INSTRUCTIONS);
  assertEquals(b.input, buildInput({ ...DI, body: "Pague em https://exemplo.pt/pay" }));
  assertEquals(b.text.format, { type: "json_schema", name: "email_summary", strict: true, schema: JSON.parse(JSON.stringify(SUMMARY_SCHEMA)) });
  assertEquals(b.text.format.schema.properties.points, { type: "array", minItems: 3, maxItems: 3, items: { type: "string" } });
});

Deno.test("AC-11: generateSummary maps transport failures", async () => {
  const st = (s: number) => generateSummary(DI, { apiKey: "K", model: "m", maxOutputTokens: 1 }, (() => Promise.resolve(new Response("{}", { status: s }))) as unknown as typeof fetch);
  assertEquals(await st(503), { ok: false, reason: "ai_unavailable", retryable: true });
  assertEquals(await st(401), { ok: false, reason: "ai_unavailable", retryable: false });
});
