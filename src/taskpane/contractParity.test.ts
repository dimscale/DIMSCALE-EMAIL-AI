/* global process */
// QA (independent) for F-003 amendment (g): the function (Deno, supabase/) and the pane (Vitest) share no module, so this
// test reads the function source and checks both sides against the same frozen list. Mocked fetch only.
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { analyze, WHY_CODES } from "./analyzeClient";

const fnSource = readFileSync(resolve(process.cwd(), "supabase/functions/_shared/ai.ts"), "utf8");
const payload = {
  action: "summarize" as const,
  subject: "S",
  fromName: "Ann",
  fromDomain: "x.pt",
  body: "b",
};
const reply = (status: number, body: unknown) =>
  vi.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  }) as any;
const unavailable = (retryable: boolean, why: string) => ({
  ok: false,
  error: "ai_unavailable",
  retryable,
  why,
});
const rejected = { kind: "error", code: "rejected" };
const unavailableCode = { kind: "error", code: "unavailable" };

describe("function and pane agree on the why codes (F-003 amendment (g))", () => {
  it("the function's WHY_CODES are the frozen list, and the pane adds only badbody", () => {
    const list = /export const WHY_CODES = \[([^\]]*)\]/.exec(fnSource);
    const fnWhy = [...(list?.[1] ?? "").matchAll(/"([^"]+)"/g)].map((m) => m[1]);
    expect(fnWhy).toEqual([
      "empty",
      "toolong",
      "newlink",
      "markup",
      "schema",
      "incomplete",
      "link",
    ]);
    expect([...WHY_CODES]).toEqual([...fnWhy, "badbody"]);
  });

  it("the function builds http_<status> on both 503 paths and sends timeout, network and badbody", () => {
    expect(fnSource.split("`http_${res.status}`").length - 1).toBe(2);
    expect(fnSource).toContain('"timeout" : "network"');
    expect(fnSource).toContain('why: "badbody"');
  });

  it.each([
    [
      "http 400, not retryable",
      503,
      unavailable(false, "http_400"),
      { ...rejected, why: "http_400" },
    ],
    [
      "http 401, not retryable",
      503,
      unavailable(false, "http_401"),
      { ...rejected, why: "http_401" },
    ],
    [
      "http 429, retryable",
      503,
      unavailable(true, "http_429"),
      { ...unavailableCode, why: "http_429" },
    ],
    [
      "http 500, retryable",
      503,
      unavailable(true, "http_500"),
      { ...unavailableCode, why: "http_500" },
    ],
    [
      "timeout, retryable",
      503,
      unavailable(true, "timeout"),
      { ...unavailableCode, why: "timeout" },
    ],
    [
      "network, retryable",
      503,
      unavailable(true, "network"),
      { ...unavailableCode, why: "network" },
    ],
    ["timeout, not retryable (never sent): dropped", 503, unavailable(false, "timeout"), rejected],
    [
      "unreadable 2xx body",
      502,
      { ok: false, error: "ai_output_invalid", why: "badbody" },
      { kind: "error", code: "invalid", why: "badbody" },
    ],
    [
      "link in the reply",
      502,
      { ok: false, error: "ai_output_invalid", why: "link" },
      { kind: "error", code: "invalid", why: "link" },
    ],
  ])(
    "%s: the pane shows the frozen code and keeps the why",
    async (_name, status, body, expected) => {
      expect(await analyze("t", payload, reply(status, body))).toEqual(expected);
    }
  );
});
