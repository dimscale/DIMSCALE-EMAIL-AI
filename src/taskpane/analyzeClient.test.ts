/* global AbortController, AbortSignal */
import { describe, expect, it, vi } from "vitest";
import { analyze } from "./analyzeClient";
import { ANALYZE_URL } from "./config";

const payload = {
  action: "summarize" as const,
  subject: "S",
  fromName: "Ann",
  fromDomain: "x.pt",
  body: "b",
};
const resp = (status: number, body: unknown) =>
  vi.fn().mockResolvedValue({ ok: status >= 200 && status < 300, status, json: async () => body });
const code = async (status: number, body: unknown) => {
  const r = await analyze("t", payload, resp(status, body) as any);
  return r.kind === "error" ? r.code : r.kind;
};

describe("analyzeClient (F-003/AC-18, AC-21)", () => {
  it("POSTs the exact payload with Bearer and JSON headers", async () => {
    const f = resp(200, { ok: true, action: "reply", reply: "Olá" });
    const reply = {
      ...payload,
      action: "reply" as const,
      language: "pt" as const,
      tone: "formal" as const,
    };
    const r = await analyze("tok", reply, f as any);
    expect(r).toEqual({ kind: "reply", reply: "Olá" });
    const [url, init] = f.mock.calls[0];
    expect(url).toBe(ANALYZE_URL);
    expect(init.method).toBe("POST");
    expect(init.headers).toEqual({
      Authorization: "Bearer tok",
      "Content-Type": "application/json",
    });
    expect(JSON.parse(init.body)).toEqual(reply);
  });
  it("parses a summary", async () => {
    const summary = {
      points: ["a", "b", "c"],
      type: "request",
      priority: "high",
      tone: "neutral",
      deadline: null,
      requestedAction: "x",
    };
    expect(
      await analyze("t", payload, resp(200, { ok: true, action: "summarize", summary }) as any)
    ).toEqual({ kind: "summary", summary });
  });
  it("maps every error row", async () => {
    expect(await code(401, { ok: false, error: "unauthorized" })).toBe("unauthorized");
    expect(await code(422, { ok: false, error: "payload_invalid" })).toBe("payload");
    expect(await code(429, { ok: false, error: "cap_reached" })).toBe("cap");
    expect(await code(503, { ok: false, error: "disabled" })).toBe("disabled");
    expect(await code(503, { ok: false, error: "ai_not_configured" })).toBe("not_configured");
    expect(await code(503, { ok: false, error: "ai_unavailable", retryable: true })).toBe(
      "unavailable"
    );
    expect(await code(503, { ok: false, error: "ai_unavailable", retryable: false })).toBe(
      "invalid"
    );
    expect(await code(502, { ok: false, error: "ai_output_invalid" })).toBe("invalid");
    expect(await code(500, { ok: false, error: "server_error" })).toBe("other");
    expect(await code(503, { ok: false, error: "server_error" })).toBe("other");
    expect(await code(429, { ok: false, error: "other" })).toBe("other");
    expect(await code(200, { ok: true, action: "summarize", summary: { points: [1] } })).toBe(
      "other"
    );
    expect(await code(502, "not json")).toBe("other");
  });
  it("maps a network error to unavailable and aborts after 40 s", async () => {
    expect(await analyze("t", payload, vi.fn().mockRejectedValue(new Error("x")) as any)).toEqual({
      kind: "error",
      code: "unavailable",
    });
    vi.useFakeTimers();
    const f = vi.fn(
      (_u: string, init: any) =>
        new Promise((_res, rej) =>
          init.signal.addEventListener("abort", () => rej(new Error("abort")))
        )
    );
    const p = analyze("t", payload, f as any);
    await vi.advanceTimersByTimeAsync(39999);
    expect(f.mock.calls[0][1].signal.aborted).toBe(false);
    await vi.advanceTimersByTimeAsync(1);
    expect(await p).toEqual({ kind: "error", code: "unavailable" });
    vi.useRealTimers();
  });
});

describe("analyzeClient translate and abort (F-003/AC-46)", () => {
  const tr = {
    action: "translate" as const,
    subject: "S",
    fromName: "",
    fromDomain: "",
    body: "b",
    targetLanguage: "en" as const,
  };
  it("maps a translate response to {subject, body} and sends the payload as is", async () => {
    const f = resp(200, { ok: true, action: "translate", translation: { subject: "", body: "x" } });
    expect(await analyze("t", tr, f as any)).toEqual({
      kind: "translation",
      translation: { subject: "", body: "x" },
    });
    expect(JSON.parse(f.mock.calls[0][1].body)).toEqual(tr);
  });
  it("accepts all three targetLanguage values and sends each one as is", async () => {
    for (const targetLanguage of ["auto", "pt", "en"] as const) {
      const f = resp(200, {
        ok: true,
        action: "translate",
        translation: { subject: "", body: "x" },
      });
      const p = { ...tr, targetLanguage };
      expect(await analyze("t", p, f as any)).toEqual({
        kind: "translation",
        translation: { subject: "", body: "x" },
      });
      expect(JSON.parse(f.mock.calls[0][1].body)).toEqual(p);
    }
  });
  it("rejects a 200 without a valid translation as the generic error", async () => {
    for (const body of [
      { ok: true, action: "translate" },
      { ok: true, action: "translate", translation: { subject: "s" } },
      { ok: true, action: "translate", translation: { subject: 1, body: "b" } },
      { ok: true, action: "reply", translation: { subject: "s", body: "b" } },
    ])
      expect(await analyze("t", tr, resp(200, body) as any)).toEqual({
        kind: "error",
        code: "other",
      });
  });
  it("keeps the same error mapping for translate", async () => {
    expect(await analyze("t", tr, resp(429, { ok: false, error: "cap_reached" }) as any)).toEqual({
      kind: "error",
      code: "cap",
    });
  });
  it("an abort resolves to aborted, not an error; an already aborted signal makes no request", async () => {
    const ctl = new AbortController();
    const f = vi.fn(
      (_u: string, init: { signal: AbortSignal }) =>
        new Promise((_res, rej) => init.signal.addEventListener("abort", () => rej(new Error("x"))))
    );
    const p = analyze("t", payload, f as any, ctl.signal);
    ctl.abort();
    expect(await p).toEqual({ kind: "aborted" });
    expect(f.mock.calls[0][1].signal.aborted).toBe(true);
    const g = resp(200, {});
    expect(await analyze("t", payload, g as any, ctl.signal)).toEqual({ kind: "aborted" });
    expect(g).not.toHaveBeenCalled();
  });
});
