/* global globalThis */
import { afterEach, describe, expect, it, vi } from "vitest";
import { readAutoSummary, saveAutoSummary } from "./prefs";

// F-003/AC-42: roamingSettings preference. Office is mocked.
const mock = (rs: unknown) => {
  (globalThis as any).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: { roamingSettings: rs },
  };
};
afterEach(() => {
  delete (globalThis as any).Office;
});

describe("prefs (F-003/AC-42)", () => {
  it("reads true only for a stored boolean true", () => {
    for (const [v, want] of [
      [true, true],
      [false, false],
      ["true", false],
      [1, false],
      [null, false],
      [undefined, false],
    ] as const) {
      mock({ get: vi.fn(() => v) });
      expect(readAutoSummary()).toBe(want);
    }
  });
  it("reads off for a throwing get and for no roamingSettings", () => {
    mock({
      get: () => {
        throw new Error("x");
      },
    });
    expect(readAutoSummary()).toBe(false);
    mock(undefined);
    expect(readAutoSummary()).toBe(false);
  });
  it("saves with set then one saveAsync and resolves true on success", async () => {
    const rs = { set: vi.fn(), saveAsync: vi.fn((cb: any) => cb({ status: "succeeded" })) };
    mock(rs);
    expect(await saveAutoSummary(true)).toBe(true);
    expect(rs.set).toHaveBeenCalledWith("autoSummary", true);
    expect(rs.saveAsync).toHaveBeenCalledTimes(1);
  });
  it("resolves false, never rejects, on a failed status, a throw or no API", async () => {
    mock({ set: vi.fn(), saveAsync: (cb: any) => cb({ status: "failed" }) });
    expect(await saveAutoSummary(false)).toBe(false);
    mock({
      set: vi.fn(),
      saveAsync: () => {
        throw new Error("x");
      },
    });
    expect(await saveAutoSummary(true)).toBe(false);
    mock(undefined);
    expect(await saveAutoSummary(true)).toBe(false);
  });
});
