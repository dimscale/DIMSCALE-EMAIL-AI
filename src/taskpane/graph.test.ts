import { describe, expect, it, vi } from "vitest";
import { filterRecentSenders, isNoReplySender, fetchRecentSenders } from "./graph";

describe("graph helpers (F-013/AC-33)", () => {
  it("detects no-reply senders by local part", () => {
    expect(isNoReplySender("noreply@company.com")).toBe(true);
    expect(isNoReplySender("NO-REPLY@service.org")).toBe(true);
    expect(isNoReplySender("donotreply@domain.pt")).toBe(true);
    expect(isNoReplySender("do-not-reply@info.com")).toBe(true);
    expect(isNoReplySender("mailer-daemon@mx.google.com")).toBe(true);
    expect(isNoReplySender("postmaster@exchange.dimscale.com")).toBe(true);

    expect(isNoReplySender("ana.silva@dimscale.com")).toBe(false);
    expect(isNoReplySender("pedro@exemplo.pt")).toBe(false);
    expect(isNoReplySender("invalid-email")).toBe(true);
  });

  it("filters and deduplicates senders, excluding ownAddress and no-reply", () => {
    const raw = [
      { from: { emailAddress: { address: "ana@cliente.pt", name: "Ana Silva" } } },
      { from: { emailAddress: { address: "ANA@cliente.pt", name: "Ana S." } } }, // duplicate
      { from: { emailAddress: { address: "noreply@sistema.pt", name: "Sistema" } } }, // no-reply
      { from: { emailAddress: { address: "eu@dimscale.com", name: "Eu Próprio" } } }, // own address
      { from: { emailAddress: { address: "rui@empresa.com", name: "Rui Costa" } } },
    ];

    const result = filterRecentSenders(raw, "eu@dimscale.com");
    expect(result).toHaveLength(2);
    expect(result[0]).toEqual({ address: "ana@cliente.pt", name: "Ana Silva" });
    expect(result[1]).toEqual({ address: "rui@empresa.com", name: "Rui Costa" });
  });

  it("caps candidates at 50", () => {
    const raw = Array.from({ length: 65 }, (_, i) => ({
      from: { emailAddress: { address: `user${i}@teste.pt`, name: `User ${i}` } },
    }));

    const result = filterRecentSenders(raw);
    expect(result).toHaveLength(50);
  });

  it("fetches recent senders and follows @odata.nextLink", async () => {
    const mockFetch = vi
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          value: [{ from: { emailAddress: { address: "first@teste.pt", name: "First" } } }],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/page2",
        }),
      } as any)
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          value: [{ from: { emailAddress: { address: "second@teste.pt", name: "Second" } } }],
        }),
      } as any);

    const candidates = await fetchRecentSenders("mock_token", "me@dimscale.com", mockFetch as any);
    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(candidates).toHaveLength(2);
    expect(candidates[0].address).toBe("first@teste.pt");
    expect(candidates[1].address).toBe("second@teste.pt");
  });
});
