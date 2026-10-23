/* global globalThis */
import { describe, expect, it, vi } from "vitest";
import { webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { hexLuminance, resolveTheme } from "./theme";

describe("theme helpers (F-012/AC-7)", () => {
  it("computes luminance and rejects invalid hex", () => {
    expect(hexLuminance("#000000")).toBe(0);
    expect(hexLuminance("#ffffff")).toBeCloseTo(1, 5);
    for (const bad of [undefined, "", "red", "#fff", "#12345g"])
      expect(hexLuminance(bad)).toBeNull();
  });
  it("dark Office theme gives dark and overrides background and foreground", () => {
    const r = resolveTheme({
      office: { bodyBackgroundColor: "#292929", bodyForegroundColor: "#FFFFFF" },
    });
    expect(r.kind).toBe("dark");
    expect(r.theme.colorNeutralBackground1).toBe("#292929");
    expect(r.theme.colorNeutralForeground1).toBe("#FFFFFF");
    expect(r.theme.colorBrandBackground).toBe(webDarkTheme.colorBrandBackground);
  });
  it("light Office theme gives light", () => {
    const r = resolveTheme({
      office: { bodyBackgroundColor: "#F3F2F1", bodyForegroundColor: "#201F1E" },
      prefersDark: true,
    });
    expect(r.kind).toBe("light");
    expect(r.theme.colorNeutralBackground1).toBe("#F3F2F1");
  });
  it("falls back to prefers-color-scheme, then light, when the Office colour is missing or invalid", () => {
    expect(resolveTheme({ prefersDark: true }).kind).toBe("dark");
    expect(resolveTheme({ office: { bodyBackgroundColor: "nope" }, prefersDark: true }).kind).toBe(
      "dark"
    );
    expect(resolveTheme({}).theme).toBe(webLightTheme);
  });
  it("forced colors wins", () => {
    expect(resolveTheme({ forcedColors: true, prefersDark: true }).kind).toBe("highContrast");
  });
});

describe("useOfficeTheme (F-012/AC-7)", () => {
  it("registers OfficeThemeChanged, updates on it, and unregisters on unmount; works with a bare mock", async () => {
    const React = await import("react");
    const { act, render, cleanup } = await import("@testing-library/react");
    const { useOfficeTheme } = await import("./theme");
    const probe = { kind: "" };
    const P = () => {
      probe.kind = useOfficeTheme().kind;
      return null;
    };
    (globalThis as any).Office = { context: {} }; // no requirements, no officeTheme
    render(React.createElement(P));
    expect(probe.kind).toBe("light");
    cleanup();

    let handler: ((a: unknown) => void) | undefined;
    const addHandlerAsync = vi.fn((_e: string, h: (a: unknown) => void) => (handler = h));
    const removeHandlerAsync = vi.fn();
    (globalThis as any).Office = {
      EventType: { OfficeThemeChanged: "officeThemeChanged" },
      context: {
        requirements: { isSetSupported: () => true },
        officeTheme: { bodyBackgroundColor: "#FFFFFF", bodyForegroundColor: "#000000" },
        mailbox: { addHandlerAsync, removeHandlerAsync },
      },
    };
    const { unmount } = render(React.createElement(P));
    expect(probe.kind).toBe("light");
    expect(addHandlerAsync.mock.calls[0][0]).toBe("officeThemeChanged");
    act(() =>
      handler?.({ officeTheme: { bodyBackgroundColor: "#141414", bodyForegroundColor: "#FFFFFF" } })
    );
    expect(probe.kind).toBe("dark");
    unmount();
    expect(removeHandlerAsync).toHaveBeenCalledTimes(1);
  });
});
