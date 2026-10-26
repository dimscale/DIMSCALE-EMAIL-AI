import { describe, expect, it, vi } from "vitest";
import { fetchSettings, saveSettings, deleteSettings } from "./settingsClient";

describe("settingsClient (F-013/AC-35, AC-38)", () => {
  it("fetches settings successfully", async () => {
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        enabled: true,
        status: "active",
        contacts: ["ana@cliente.pt", "rui@empresa.com"],
      }),
    } as any);

    const res = await fetchSettings("token_123", mockFetch as any);
    expect(res.kind).toBe("ok");
    if (res.kind === "ok") {
      expect(res.data.enabled).toBe(true);
      expect(res.data.status).toBe("active");
      expect(res.data.contacts).toEqual(["ana@cliente.pt", "rui@empresa.com"]);
    }
  });

  it("handles 401 unauthorized on fetchSettings", async () => {
    const mockFetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
    } as any);

    const res = await fetchSettings("bad_token", mockFetch as any);
    expect(res.kind).toBe("rejected");
  });

  it("saves settings with PUT and payload", async () => {
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      json: async () => ({
        enabled: true,
        status: "pending",
        contacts: ["test@example.com"],
      }),
    } as any);

    const payload = {
      enabled: true,
      ownAddress: "me@dimscale.com",
      contacts: ["test@example.com"],
    };

    const res = await saveSettings("token_123", payload, mockFetch as any);
    expect(mockFetch).toHaveBeenCalledWith(
      expect.stringContaining("/auto-reply-settings"),
      expect.objectContaining({
        method: "PUT",
        body: JSON.stringify(payload),
      })
    );
    expect(res.kind).toBe("ok");
  });

  it("deletes settings with DELETE", async () => {
    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
    } as any);

    const res = await deleteSettings("token_123", mockFetch as any);
    expect(mockFetch).toHaveBeenCalledWith(
      expect.stringContaining("/auto-reply-settings"),
      expect.objectContaining({ method: "DELETE" })
    );
    expect(res.ok).toBe(true);
  });
});
