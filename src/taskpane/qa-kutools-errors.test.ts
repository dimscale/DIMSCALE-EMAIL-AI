/* global AbortSignal */
import { describe, expect, it, vi } from "vitest";
import { analyze, AnalyzePayload } from "./analyzeClient";

// QA (independent): analyzeClient error table and malformed input (F-003/AC-21, AC-46). Mocked fetch only.
const TOKEN = "eyJ-SECRET-TOKEN-VALUE";
const base = { subject: "S", fromName: "Ann", fromDomain: "x.pt", body: "b" };
const summarize: AnalyzePayload = { action: "summarize", ...base };
const reply: AnalyzePayload = { action: "reply", ...base, language: "match", tone: "none" };
const translate: AnalyzePayload = {
  action: "translate",
  subject: "S",
  fromName: "",
  fromDomain: "",
  body: "b",
  targetLanguage: "auto",
};

/* A response whose json() either returns the body or rejects (malformed JSON). */
const resp = (status: number, body: unknown, badJson = false) =>
  vi.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: badJson
      ? async () => {
          throw new SyntaxError("Unexpected token < in JSON at position 0");
        }
      : async () => body,
  });
const code = async (p: AnalyzePayload, status: number, body: unknown, badJson = false) => {
  const r = await analyze(TOKEN, p, resp(status, body, badJson) as any);
  return r.kind === "error" ? r.code : r.kind;
};

describe("error table, every status and malformed bodies (F-003/AC-21)", () => {
  it("a malformed JSON body maps to the generic error for 200, 401, 422, 429, 500, 502 and 503", async () => {
    for (const status of [200, 401, 422, 429, 500, 502, 503])
      expect(await code(summarize, status, null, true)).toBe(
        status === 401 ? "unauthorized" : status === 422 ? "payload" : "other"
      );
  });

  it("a 401 with a malformed body is still unauthorized (the status decides, not the body)", async () => {
    expect(await code(summarize, 401, null, true)).toBe("unauthorized");
  });

  it("503 ai_unavailable: retryable must be the boolean true; the string 'true' or a missing flag is not retryable", async () => {
    expect(await code(reply, 503, { ok: false, error: "ai_unavailable", retryable: "true" })).toBe(
      "invalid"
    );
    expect(await code(reply, 503, { ok: false, error: "ai_unavailable" })).toBe("invalid");
    expect(await code(reply, 503, { ok: false, error: "ai_unavailable", retryable: true })).toBe(
      "unavailable"
    );
  });

  it("429 with any error other than cap_reached is the generic error, not the cap text", async () => {
    expect(await code(summarize, 429, { ok: false, error: "rate_limited" })).toBe("other");
    expect(await code(summarize, 429, {})).toBe("other");
  });

  it("503 disabled and ai_not_configured only match on status 503; the same error on 500 is generic", async () => {
    expect(await code(summarize, 500, { ok: false, error: "disabled" })).toBe("other");
    expect(await code(summarize, 500, { ok: false, error: "ai_not_configured" })).toBe("other");
  });

  it("a 200 with ok false is the generic error", async () => {
    expect(await code(summarize, 200, { ok: false, action: "summarize", summary: {} })).toBe(
      "other"
    );
  });

  it("OBSERVATION: a 200 whose action differs from the request is accepted as the other card's result (analyze does not compare actions)", async () => {
    expect(
      await code(reply, 200, {
        ok: true,
        action: "summarize",
        summary: { points: ["a", "b", "c"] },
      })
    ).toBe("summary");
    expect(
      await code(reply, 200, {
        ok: true,
        action: "translate",
        translation: { subject: "", body: "x" },
      })
    ).toBe("translation");
  });

  it("a 200 reply with a non-string reply, and a translation with null, are the generic error", async () => {
    expect(await code(reply, 200, { ok: true, action: "reply", reply: 42 })).toBe("other");
    expect(await code(translate, 200, { ok: true, action: "translate", translation: null })).toBe(
      "other"
    );
    expect(
      await code(translate, 200, {
        ok: true,
        action: "translate",
        translation: { subject: "s", body: null },
      })
    ).toBe("other");
  });

  it("summary fields of the wrong type are coerced safely: numbers become '' or null, never a crash", async () => {
    const r = await analyze(
      TOKEN,
      summarize,
      resp(200, {
        ok: true,
        action: "summarize",
        summary: {
          points: ["a", "b", "c"],
          type: 5,
          priority: null,
          tone: {},
          deadline: 7,
          requestedAction: "",
        },
      }) as any
    );
    expect(r).toEqual({
      kind: "summary",
      summary: {
        points: ["a", "b", "c"],
        type: "",
        priority: "",
        tone: "",
        deadline: null,
        requestedAction: "",
      },
    });
  });

  it("the client does not check the point count or emptiness (server rule, F-003 AC-9): recorded as an observation", async () => {
    const r = await analyze(
      TOKEN,
      summarize,
      resp(200, { ok: true, action: "summarize", summary: { points: [] } }) as any
    );
    expect(r.kind).toBe("summary");
  });
});

describe("no token or response content in error results (F-003/AC-12, AC-25)", () => {
  it("no error or result object carries the access token or the error body text", async () => {
    const cases: [AnalyzePayload, number, unknown, boolean][] = [
      [summarize, 401, { ok: false, error: "unauthorized", detail: TOKEN }, false],
      [reply, 503, { ok: false, error: "ai_unavailable", retryable: true, detail: TOKEN }, false],
      [translate, 502, null, true],
      [summarize, 500, { ok: false, error: "server_error", detail: TOKEN }, false],
    ];
    for (const [p, status, body, bad] of cases) {
      const r = await analyze(TOKEN, p, resp(status, body, bad) as any);
      expect(JSON.stringify(r)).not.toContain(TOKEN);
    }
    const net = await analyze(
      TOKEN,
      summarize,
      vi.fn().mockRejectedValue(new Error(`net ${TOKEN}`)) as any
    );
    expect(JSON.stringify(net)).not.toContain(TOKEN);
  });

  it("the authorization header is the only place the token is used, and the request body never carries it", async () => {
    const f = resp(200, { ok: true, action: "reply", reply: "r" });
    await analyze(TOKEN, reply, f as any);
    const [, init] = f.mock.calls[0];
    expect(init.headers.Authorization).toBe(`Bearer ${TOKEN}`);
    expect(init.body).not.toContain(TOKEN);
  });
});

describe("network failure and timeout (F-003/AC-21)", () => {
  it("a rejected fetch maps to unavailable (retryable), never to an error text containing the message", async () => {
    const r = await analyze(
      TOKEN,
      summarize,
      vi.fn().mockRejectedValue(new TypeError("Failed to fetch")) as any
    );
    expect(r).toEqual({ kind: "error", code: "unavailable" });
  });

  it("an aborted fetch that was not asked to abort by the caller is 'unavailable', not 'aborted'", async () => {
    vi.useFakeTimers();
    try {
      const f = vi.fn(
        (_u: string, init: { signal: AbortSignal }) =>
          new Promise((_res, rej) =>
            init.signal.addEventListener("abort", () => rej(new Error("AbortError")))
          )
      );
      const p = analyze(TOKEN, summarize, f as any);
      await vi.advanceTimersByTimeAsync(40000);
      expect(await p).toEqual({ kind: "error", code: "unavailable" });
    } finally {
      vi.useRealTimers();
    }
  });
});
