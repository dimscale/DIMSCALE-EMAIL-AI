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
    const r = await analyze("tok", { ...payload, action: "reply" }, f as any);
    expect(r).toEqual({ kind: "reply", reply: "Olá" });
    const [url, init] = f.mock.calls[0];
    expect(url).toBe(ANALYZE_URL);
    expect(init.method).toBe("POST");
    expect(init.headers).toEqual({
      Authorization: "Bearer tok",
      "Content-Type": "application/json",
    });
    expect(JSON.parse(init.body)).toEqual({ ...payload, action: "reply" });
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
