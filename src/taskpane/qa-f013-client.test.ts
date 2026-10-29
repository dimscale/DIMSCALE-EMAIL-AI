/* global RequestInit */
// QA (independent, F-013 step 10b): Graph candidate listing and settings client, mocked fetch only.
import { afterEach, describe, expect, it, vi } from "vitest";
import { fetchRecentSenders, filterRecentSenders } from "./graph";
import { deleteSettings, fetchSettings, saveSettings } from "./settingsClient";
import { SETTINGS_URL } from "./config";

const GRAPH_FIRST_PAGE = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages";
const GRAPH_NEXT = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=100";

const ok = (value: unknown[], nextLink?: string) =>
  ({
    ok: true,
    status: 200,
    json: async () => ({ value, ...(nextLink ? { "@odata.nextLink": nextLink } : {}) }),
  }) as any;

const status = (code: number) => ({ ok: false, status: code, json: async () => ({}) }) as any;

const msg = (address: string, name = "") => ({ from: { emailAddress: { address, name } } });

const many = (n: number, prefix = "user") =>
  Array.from({ length: n }, (_, i) => msg(`${prefix}${i}@teste.pt`, `User ${i}`));

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

describe("F-013/AC-32 candidate request (mocked fetch)", () => {
  it("sends exactly the specified Inbox query with the bearer token and no body", async () => {
    const f = vi.fn().mockResolvedValue(ok([]));
    await fetchRecentSenders("tok_123", "me@dimscale.com", f as any);

    expect(f).toHaveBeenCalledTimes(1);
    const [url, init] = f.mock.calls[0] as [string, RequestInit];
    expect(url.startsWith(`${GRAPH_FIRST_PAGE}?$select=from&$filter=receivedDateTime ge `)).toBe(
      true
    );
    expect(url.endsWith("&$orderby=receivedDateTime desc&$top=100")).toBe(true);
    expect(init.method).toBe("GET");
    expect(init.headers).toEqual({ Authorization: "Bearer tok_123" });
    expect(init.body).toBeUndefined();
  });

  it("uses a 30-day receivedDateTime window in the filter", async () => {
    const f = vi.fn().mockResolvedValue(ok([]));
    const before = Date.now();
    await fetchRecentSenders("tok", undefined, f as any);
    const url = String(f.mock.calls[0][0]);
    const iso = /receivedDateTime ge (\S+?)&/.exec(url)?.[1];
    expect(iso).toBeTruthy();
    const ageMs = before - Date.parse(iso!);
    expect(ageMs).toBeGreaterThanOrEqual(30 * 24 * 3600 * 1000 - 5000);
    expect(ageMs).toBeLessThanOrEqual(30 * 24 * 3600 * 1000 + 5000);
  });

  it("rejects on Graph 429 (no partial list is returned)", async () => {
    const f = vi.fn().mockResolvedValue(status(429));
    await expect(fetchRecentSenders("tok", "me@x.pt", f as any)).rejects.toThrow("429");
  });

  it("rejects on Graph 500", async () => {
    const f = vi.fn().mockResolvedValue(status(500));
    await expect(fetchRecentSenders("tok", "me@x.pt", f as any)).rejects.toThrow("500");
  });

  it("rejects when a later page fails, even after a good first page", async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(ok([msg("a@teste.pt")], GRAPH_NEXT))
      .mockResolvedValueOnce(status(503));
    await expect(fetchRecentSenders("tok", undefined, f as any)).rejects.toThrow("503");
  });

  it("rejects on network failure", async () => {
    const f = vi.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    await expect(fetchRecentSenders("tok", "me@x.pt", f as any)).rejects.toThrow();
  });

  it("aborts a hung page after 10 s and rejects", async () => {
    vi.useFakeTimers();
    const f = vi.fn().mockImplementation(
      (_url: string, init: RequestInit) =>
        new Promise((_, reject) => {
          init.signal?.addEventListener("abort", () => reject(new Error("aborted")));
        })
    );
    const p = fetchRecentSenders("tok", undefined, f as any);
    const assertion = expect(p).rejects.toThrow("aborted");
    await vi.advanceTimersByTimeAsync(10000);
    await assertion;
  });

  it("treats a page without a value array as empty, without throwing", async () => {
    const f = vi.fn().mockResolvedValue({ ok: true, status: 200, json: async () => ({}) } as any);
    await expect(fetchRecentSenders("tok", "me@x.pt", f as any)).resolves.toEqual([]);
  });

  it("an empty page ends paging even when a nextLink is present", async () => {
    const f = vi.fn().mockResolvedValueOnce(ok([], GRAPH_NEXT));
    await expect(fetchRecentSenders("tok", undefined, f as any)).resolves.toEqual([]);
    expect(f).toHaveBeenCalledTimes(1);
  });

  it("follows a nextLink on graph.microsoft.com (any letter case of the host)", async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(ok([msg("a@teste.pt")], GRAPH_NEXT))
      .mockResolvedValueOnce(ok([msg("b@teste.pt")]));
    const res = await fetchRecentSenders("tok", undefined, f as any);
    expect(f.mock.calls[1][0]).toBe(GRAPH_NEXT);
    expect(res.map((c) => c.address)).toEqual(["a@teste.pt", "b@teste.pt"]);

    const g = vi
      .fn()
      .mockResolvedValueOnce(ok([msg("a@teste.pt")], "https://GRAPH.microsoft.com/v1.0/next"))
      .mockResolvedValueOnce(ok([msg("b@teste.pt")]));
    await fetchRecentSenders("tok", undefined, g as any);
    expect(g).toHaveBeenCalledTimes(2);
  });

  // Fixed in 2837c95 (graph.ts nextLink host check). Foreign or malformed nextLinks must not receive the token.
  it.each([
    ["another host", "https://evil.example.com/collect?x=1"],
    ["a host that starts with graph.microsoft.com", "https://graph.microsoft.com.evil.example/x"],
    ["userinfo trick", "https://graph.microsoft.com@evil.example/x"],
    ["plain http on graph", "http://graph.microsoft.com/v1.0/next"],
    ["relative URL", "/v1.0/next"],
    ["not a URL", "not a url"],
  ])("does NOT follow a nextLink to %s (token must not leave Graph)", async (_label, next) => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(ok([msg("a@teste.pt")], next))
      .mockResolvedValue(ok([msg("b@teste.pt")]));
    const res = await fetchRecentSenders("tok", undefined, f as any);
    expect(f).toHaveBeenCalledTimes(1);
    expect(f.mock.calls.every((c) => String(c[0]).startsWith(GRAPH_FIRST_PAGE))).toBe(true);
    expect(res.map((c) => c.address)).toEqual(["a@teste.pt"]);
  });

  it("stops at 300 messages even when nextLink keeps going (3 pages of 100)", async () => {
    const f = vi
      .fn()
      .mockImplementation(async () => ok(many(100, `p${f.mock.calls.length}_`), GRAPH_NEXT));
    const res = await fetchRecentSenders("tok", undefined, f as any);
    expect(f).toHaveBeenCalledTimes(3);
    expect(res).toHaveLength(50);
  });

  it("stops after 5 pages in any case (page cap)", async () => {
    const f = vi
      .fn()
      .mockImplementation(async () => ok([msg(`p${f.mock.calls.length}@teste.pt`)], GRAPH_NEXT));
    await fetchRecentSenders("tok", undefined, f as any);
    expect(f).toHaveBeenCalledTimes(5);
  });
});

describe("F-013/AC-33 candidate filtering", () => {
  it("excludes the own address case-insensitively and with whitespace", () => {
    const raw = [
      msg("thiago.santos@dimscale.com"),
      msg("THIAGO.SANTOS@DIMSCALE.com "),
      msg("  Thiago.Santos@Dimscale.COM"),
    ];
    expect(filterRecentSenders(raw, "  Thiago.Santos@DIMSCALE.com ")).toEqual([]);
  });

  it("does not exclude a different address at the same domain", () => {
    const raw = [msg("thiago.santos@dimscale.com"), msg("outro@dimscale.com")];
    expect(filterRecentSenders(raw, "thiago.santos@dimscale.com").map((c) => c.address)).toEqual([
      "outro@dimscale.com",
    ]);
  });

  it("excludes no-reply local parts case-insensitively", () => {
    const raw = [
      msg("NoReply@x.pt"),
      msg("Do-Not-Reply@x.pt"),
      msg("MAILER-DAEMON@x.pt"),
      msg("ana@x.pt"),
    ];
    expect(filterRecentSenders(raw).map((c) => c.address)).toEqual(["ana@x.pt"]);
  });

  it("dedupes uppercase and whitespace variants into one lower-cased candidate, keeping the first name", () => {
    const raw = [msg("  Ana@Cliente.PT ", "Ana Silva"), msg("ana@cliente.pt", "Ana S.")];
    expect(filterRecentSenders(raw)).toEqual([{ address: "ana@cliente.pt", name: "Ana Silva" }]);
  });

  it("skips messages without a from address", () => {
    const raw = [{}, { from: {} }, { from: { emailAddress: {} } }, msg("ok@x.pt")];
    expect(filterRecentSenders(raw as any).map((c) => c.address)).toEqual(["ok@x.pt"]);
  });

  it("50-cap boundary: 49 gives 49, 50 gives 50, 51 gives 50 (the first 50 by order)", () => {
    expect(filterRecentSenders(many(49))).toHaveLength(49);
    expect(filterRecentSenders(many(50))).toHaveLength(50);
    const r51 = filterRecentSenders(many(51));
    expect(r51).toHaveLength(50);
    expect(r51[0].address).toBe("user0@teste.pt");
    expect(r51[49].address).toBe("user49@teste.pt");
  });

  it("duplicates do not count toward the 50 cap (10 repeats first, then 50 distinct)", () => {
    const raw = [...Array.from({ length: 10 }, () => msg("dup@x.pt")), ...many(50)];
    const r = filterRecentSenders(raw);
    expect(r).toHaveLength(50);
    expect(new Set(r.map((c) => c.address)).size).toBe(50);
    expect(r[0].address).toBe("dup@x.pt");
    expect(r[49].address).toBe("user48@teste.pt");
  });
});

describe("F-013/AC-35 settings client (mocked fetch, Authorization header only)", () => {
  it("network failure on GET resolves to error and never throws", async () => {
    const f = vi.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    await expect(fetchSettings("tok", f as any)).resolves.toEqual({ kind: "error" });
  });

  it("network failure on PUT resolves to error and never throws", async () => {
    const f = vi.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    await expect(
      saveSettings("tok", { enabled: false, ownAddress: "me@x.pt", contacts: [] }, f as any)
    ).resolves.toEqual({ kind: "error" });
  });

  it("network failure on DELETE resolves to ok false", async () => {
    const f = vi.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    await expect(deleteSettings("tok", f as any)).resolves.toEqual({ ok: false });
  });

  it('server status "error" is passed through as status error (not treated as ok-off)', async () => {
    const f = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({ enabled: true, status: "error", contacts: ["ana@x.pt"] }),
    } as any);
    const res = await fetchSettings("tok", f as any);
    expect(res).toEqual({
      kind: "ok",
      data: { enabled: true, status: "error", contacts: ["ana@x.pt"] },
    });
  });

  it("422 on PUT returns error with status 422 (not ok)", async () => {
    const f = vi.fn().mockResolvedValue(status(422));
    const res = await saveSettings(
      "tok",
      { enabled: true, ownAddress: "me@x.pt", contacts: ["a@x.pt"] },
      f as any
    );
    expect(res).toEqual({ kind: "error", status: 422 });
  });

  it("500 on GET returns error with status 500", async () => {
    const f = vi.fn().mockResolvedValue(status(500));
    expect(await fetchSettings("tok", f as any)).toEqual({ kind: "error", status: 500 });
  });

  it("401 on PUT returns rejected", async () => {
    const f = vi.fn().mockResolvedValue(status(401));
    expect(
      await saveSettings("tok", { enabled: false, ownAddress: "me@x.pt", contacts: [] }, f as any)
    ).toEqual({ kind: "rejected" });
  });

  it("DELETE non-2xx returns ok false", async () => {
    const f = vi.fn().mockResolvedValue(status(500));
    expect(await deleteSettings("tok", f as any)).toEqual({ ok: false });
  });

  it("requests go to the configured settings URL with the token only in the header", async () => {
    const f = vi.fn().mockResolvedValue(status(500));
    await fetchSettings("tok_secret", f as any);
    const [url, init] = f.mock.calls[0] as [string, RequestInit];
    expect(url).toBe(SETTINGS_URL);
    expect(url).not.toContain("tok_secret");
    expect(init.headers).toEqual({ Authorization: "Bearer tok_secret" });
  });

  it("PUT carries the bearer header and Content-Type application/json (AC-35)", async () => {
    const f = vi.fn().mockResolvedValue(status(500));
    const payload = { enabled: true, ownAddress: "me@x.pt", contacts: ["a@x.pt"] };
    await saveSettings("tok_put", payload, f as any);
    const [url, init] = f.mock.calls[0] as [string, RequestInit];
    expect(url).toBe(SETTINGS_URL);
    expect(init.method).toBe("PUT");
    expect(init.headers).toEqual({
      Authorization: "Bearer tok_put",
      "Content-Type": "application/json",
    });
    expect(JSON.parse(String(init.body))).toEqual(payload);
  });

  it("aborts a hung request after 10 s and returns error", async () => {
    vi.useFakeTimers();
    const f = vi.fn().mockImplementation(
      (_url: string, init: RequestInit) =>
        new Promise((_, reject) => {
          init.signal?.addEventListener("abort", () => reject(new Error("aborted")));
        })
    );
    const p = fetchSettings("tok", f as any);
    await vi.advanceTimersByTimeAsync(10000);
    await expect(p).resolves.toEqual({ kind: "error" });
  });
});
