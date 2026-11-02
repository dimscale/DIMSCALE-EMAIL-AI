/* global globalThis */
import { afterEach, describe, expect, it, vi } from "vitest";
import { readAuto, readConsent, saveAuto } from "./prefs";

// F-003/AC-60: roamingSettings preferences, three keys plus the remembered consent. Office is mocked.
const mock = (rs: unknown) => {
  (globalThis as any).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: { roamingSettings: rs },
  };
};
const KEYS = [
  ["summarize", "autoSummary", "autoSummaryConsent"],
  ["reply", "autoReply", "autoReplyConsent"],
  ["translate", "autoTranslate", "autoTranslateConsent"],
] as const;
afterEach(() => {
  delete (globalThis as any).Office;
});

describe.each(KEYS)("prefs for %s (F-003/AC-60)", (action, key, consentKey) => {
  it("reads true only for a stored boolean true, of its own key", () => {
    for (const [v, want] of [
      [true, true],
      [false, false],
      ["true", false],
      [1, false],
      [null, false],
      [undefined, false],
    ] as const) {
      const get = vi.fn((k: string) => (k === key ? v : undefined));
      mock({ get });
      expect(readAuto(action)).toBe(want);
      expect(get).toHaveBeenCalledWith(key);
    }
    mock({ get: vi.fn((k: string) => (k === key ? undefined : true)) }); // the other keys do not count
    expect(readAuto(action)).toBe(false);
  });
  it("reads off for a throwing get and for no roamingSettings", () => {
    mock({
      get: () => {
        throw new Error("x");
      },
    });
    expect(readAuto(action)).toBe(false);
    expect(readConsent(action)).toBe(false);
    mock(undefined);
    expect(readAuto(action)).toBe(false);
    expect(readConsent(action)).toBe(false);
  });
  it("consent is its own boolean key; an earlier stored 'on' counts as consent (Q-U22); other values do not", () => {
    mock({ get: vi.fn((k: string) => (k === consentKey ? true : undefined)) });
    expect(readConsent(action)).toBe(true);
    mock({ get: vi.fn((k: string) => (k === key ? true : undefined)) });
    expect(readConsent(action)).toBe(true);
    mock({ get: vi.fn(() => "true") });
    expect(readConsent(action)).toBe(false);
  });
  it("turning on sets its own key and its consent, then one saveAsync; turning off sets only its key", async () => {
    const rs = { set: vi.fn(), saveAsync: vi.fn((cb: any) => cb({ status: "succeeded" })) };
    mock(rs);
    expect(await saveAuto(action, true)).toBe(true);
    expect(rs.set.mock.calls).toEqual([
      [key, true],
      [consentKey, true],
    ]);
    expect(rs.saveAsync).toHaveBeenCalledTimes(1);
    rs.set.mockClear();
    expect(await saveAuto(action, false)).toBe(true);
    expect(rs.set.mock.calls).toEqual([[key, false]]);
    expect(rs.saveAsync).toHaveBeenCalledTimes(2);
  });
  it("resolves false, never rejects, on a failed status, a throw or no API", async () => {
    mock({ set: vi.fn(), saveAsync: (cb: any) => cb({ status: "failed" }) });
    expect(await saveAuto(action, false)).toBe(false);
    mock({
      set: vi.fn(),
      saveAsync: () => {
        throw new Error("x");
      },
    });
    expect(await saveAuto(action, true)).toBe(false);
    mock(undefined);
    expect(await saveAuto(action, true)).toBe(false);
  });
});
