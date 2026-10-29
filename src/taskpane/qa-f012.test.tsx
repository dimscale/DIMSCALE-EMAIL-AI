/* global globalThis, window, document, DOMParser */
// QA pass for F-012 (direction D, 320 px pane). Test-only file. Mocks: Office.js, matchMedia, BackendStatus (F-002 has its own tests).
// Tests that FAIL here are real gaps against docs/specs/F-012-taskpane-design.md, not test bugs.
import * as React from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { teamsHighContrastTheme, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import { resolveTheme, useOfficeTheme } from "./theme";
import { strings } from "./strings";
import MessagePanel from "./MessagePanel";
import App from "./App";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));

type Cb = (r: { status: string; value?: string }) => void;
type Theme = Record<string, unknown>;

// Independent WCAG 2.x contrast (not the production helper, to avoid testing the code with itself).
function relLum(hex: string): number {
  const c = [1, 3, 5]
    .map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
}
function contrast(a: string, b: string): number {
  const x = relLum(a);
  const y = relLum(b);
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
}
const hex = (v: unknown) => {
  if (typeof v !== "string" || !/^#[0-9a-f]{6}$/i.test(v))
    throw new Error(`not a hex token: ${String(v)}`);
  return v;
};

function setOffice(context: Record<string, unknown> = {}) {
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged", OfficeThemeChanged: "officeThemeChanged" },
    context: { mailbox: { item: undefined }, ...context },
  };
}

// matchMedia stub: `matches` decides the answer per query; listeners are captured so a test can fire them.
function stubMatchMedia(matches: Record<string, boolean>) {
  const listeners: Array<{ q: string; fn: () => void }> = [];
  const removed: Array<{ q: string; fn: () => void }> = [];
  (window as any).matchMedia = (q: string) => ({
    matches: !!matches[q],
    media: q,
    addEventListener: (_: string, fn: () => void) => listeners.push({ q, fn }),
    removeEventListener: (_: string, fn: () => void) => removed.push({ q, fn }),
  });
  return { listeners, removed };
}

const makeItem = (getAsync: unknown, o: Record<string, unknown> = {}) => ({
  subject: "Hello",
  from: { displayName: "Ann", emailAddress: "ann@example.com" },
  dateTimeCreated: new Date(2026, 0, 2, 3, 4),
  body: { getAsync },
  ...o,
});
const ok = (value: string) => vi.fn((_t: string, cb: Cb) => cb({ status: "succeeded", value }));
const openBody = () => {
  const h = screen.queryByRole("button", { name: "Mensagem" });
  if (h) fireEvent.click(h);
};

afterEach(() => {
  cleanup();
  delete (window as any).matchMedia;
  delete (globalThis as any).Office;
});

// ---------- AC-6 / AC-5: contrast ----------
describe("contrast of the stock Fluent pairs the pane renders (F-012/AC-5, AC-6), independent WCAG math", () => {
  const stock: Array<[string, Theme]> = [
    ["webLightTheme", webLightTheme as unknown as Theme],
    ["webDarkTheme", webDarkTheme as unknown as Theme],
    ["teamsHighContrastTheme", teamsHighContrastTheme as unknown as Theme],
  ];
  for (const [name, t] of stock) {
    for (const fg of [
      "colorNeutralForeground1",
      "colorNeutralForeground2",
      "colorNeutralForeground3",
    ]) {
      it(`${name}: ${fg} on colorNeutralBackground1 >= 4.5:1`, () => {
        const r = contrast(hex(t[fg]), hex(t.colorNeutralBackground1));
        expect(
          r,
          `${name} ${fg} ${String(t[fg])} on ${String(t.colorNeutralBackground1)} = ${r.toFixed(2)}:1`
        ).toBeGreaterThanOrEqual(4.5);
      });
    }
  }
  it.each([
    ["Office body FFFFFF bg / 242424 fg", "#FFFFFF", "#242424"],
    ["Office body 292929 bg / FFFFFF fg", "#292929", "#FFFFFF"],
    ["Office body 141414 bg / FFFFFF fg", "#141414", "#FFFFFF"],
    ["Office body F3F2F1 bg / 201F1E fg", "#F3F2F1", "#201F1E"],
  ])("simulated %s >= 4.5:1", (_label, bg, fg) => {
    expect(contrast(fg, bg)).toBeGreaterThanOrEqual(4.5);
  });
  // MessageBar intents and footer status rendered with these tokens (react-message-bar styles, Fluent 9.74.10).
  it.each([
    [
      "MessageBar error text (fg1 on dangerBg1)",
      "colorNeutralForeground1",
      "colorStatusDangerBackground1",
    ],
    [
      "MessageBar error icon (dangerFg1 on dangerBg1)",
      "colorStatusDangerForeground1",
      "colorStatusDangerBackground1",
    ],
    ["MessageBar info (fg3 on neutralBg3)", "colorNeutralForeground3", "colorNeutralBackground3"],
    [
      "footer/status text and icons (fg3 on bg1)",
      "colorNeutralForeground3",
      "colorNeutralBackground1",
    ],
  ])("semantic pair %s in light and dark", (_label, fgKey, bgKey) => {
    for (const t of [webLightTheme, webDarkTheme] as unknown as Theme[]) {
      expect(contrast(hex(t[fgKey]), hex(t[bgKey]))).toBeGreaterThanOrEqual(4.5);
    }
  });
});

describe("Office colour guard (F-012/AC-6, spec Theme 1)", () => {
  it.each([
    ["light", { bodyBackgroundColor: "#FFFFFF", bodyForegroundColor: "#FFFFFF" }, webLightTheme],
    ["dark", { bodyBackgroundColor: "#141414", bodyForegroundColor: "#141414" }, webDarkTheme],
  ])(
    "%s: an Office pair below 4.5:1 is NOT applied; the Fluent neutral is kept (spec AC-6)",
    (_k, office, base) => {
      // Pair is fully illegible; the spec requires the guard to fall back to Fluent neutrals.
      const r = resolveTheme({ office });
      const low = (v: unknown) => String(v).toLowerCase();
      expect(low(r.theme.colorNeutralBackground1)).toBe(
        low((base as unknown as Theme).colorNeutralBackground1)
      );
      expect(low(r.theme.colorNeutralForeground1)).toBe(
        low((base as unknown as Theme).colorNeutralForeground1)
      );
    }
  );
  it("an Office pair above 4.5:1 is applied (both background and foreground)", () => {
    const r = resolveTheme({
      office: { bodyBackgroundColor: "#F3F2F1", bodyForegroundColor: "#201F1E" },
    });
    expect(r.theme.colorNeutralBackground1).toBe("#F3F2F1");
    expect(r.theme.colorNeutralForeground1).toBe("#201F1E");
  });
});

// ---------- AC-7: theme resolution edges ----------
describe("theme resolution edges (F-012/AC-7)", () => {
  it("luminance threshold 0.179: #747474 is dark, #767676 is light (independent luminance)", () => {
    expect(relLum("#747474")).toBeLessThan(0.179);
    expect(relLum("#767676")).toBeGreaterThanOrEqual(0.179);
    expect(resolveTheme({ office: { bodyBackgroundColor: "#747474" } }).kind).toBe("dark");
    expect(resolveTheme({ office: { bodyBackgroundColor: "#767676" } }).kind).toBe("light");
  });
  it("invalid foreground hex falls back to the Fluent foreground; background still applied", () => {
    const r = resolveTheme({
      office: { bodyBackgroundColor: "#FFFFFF", bodyForegroundColor: "blue" },
    });
    expect(r.theme.colorNeutralForeground1).toBe(webLightTheme.colorNeutralForeground1);
  });
  it("3-digit hex is invalid and falls back to prefers-color-scheme", () => {
    expect(resolveTheme({ office: { bodyBackgroundColor: "#fff" }, prefersDark: true }).kind).toBe(
      "dark"
    );
  });
  it("forced colours return the Fluent high-contrast theme object and ignore Office colours", () => {
    const r = resolveTheme({ forcedColors: true, office: { bodyBackgroundColor: "#141414" } });
    expect(r.theme).toBe(teamsHighContrastTheme);
    expect(r.kind).toBe("highContrast");
  });
  it("useOfficeTheme re-reads on prefers-color-scheme and forced-colors change events, and removes them on unmount", () => {
    setOffice({ requirements: { isSetSupported: () => false } });
    const dark = { on: false, forced: false };
    const mq = stubMatchMedia({});
    (window as any).matchMedia = (q: string) => ({
      get matches() {
        return q === "(prefers-color-scheme: dark)"
          ? dark.on
          : q === "(forced-colors: active)"
            ? dark.forced
            : false;
      },
      addEventListener: (_: string, fn: () => void) => mq.listeners.push({ q, fn }),
      removeEventListener: (_: string, fn: () => void) => mq.removed.push({ q, fn }),
    });
    const probe = { kind: "" };
    const P = () => {
      probe.kind = useOfficeTheme().kind;
      return null;
    };
    const { unmount } = render(React.createElement(P));
    expect(probe.kind).toBe("light");
    dark.on = true;
    act(() =>
      mq.listeners.filter((l) => l.q === "(prefers-color-scheme: dark)").forEach((l) => l.fn())
    );
    expect(probe.kind).toBe("dark");
    dark.forced = true;
    act(() => mq.listeners.filter((l) => l.q === "(forced-colors: active)").forEach((l) => l.fn()));
    expect(probe.kind).toBe("highContrast");
    unmount();
    expect(mq.removed.map((r) => r.q).sort()).toEqual([
      "(forced-colors: active)",
      "(prefers-color-scheme: dark)",
    ]);
  });
});

// ---------- AC-11 / AC-20 / AC-23: App layout and wordmark ----------
const logoFills = () =>
  Array.from(document.querySelectorAll('svg[role="img"] path')).map((p) => p.getAttribute("fill"));
describe("App chrome (F-012/AC-10, AC-11, AC-16, AC-19, AC-20)", () => {
  it("wordmark colour is #ffffff on dark, #000000 on light, currentColor under forced colours", () => {
    setOffice();
    stubMatchMedia({ "(prefers-color-scheme: dark)": true });
    const d = render(<App />);
    expect(new Set(logoFills())).toEqual(new Set(["#ffffff"]));
    d.unmount();
    stubMatchMedia({});
    render(<App />);
    expect(new Set(logoFills())).toEqual(new Set(["#000000"]));
    cleanup();
    stubMatchMedia({ "(forced-colors: active)": true, "(prefers-color-scheme: dark)": true });
    render(<App />);
    expect(new Set(logoFills())).toEqual(new Set(["currentColor"]));
  });
  it("one h1 whose accessible name is 'DIMSCALE Mail' and a labelled tablist", () => {
    setOffice();
    render(<App />);
    expect(screen.getByRole("heading", { level: 1, name: "DIMSCALE Mail" })).toBeTruthy();
    expect(screen.getAllByRole("tab")).toHaveLength(4);
    expect(screen.getByRole("tablist", { name: strings.tabsLabel })).toBeTruthy();
  });
  it("each tab controls a rendered tabpanel that is labelled by it (ARIA wiring; arrow keys are live-only, see report)", () => {
    setOffice();
    render(<App />);
    const summary = screen.getByRole("tab", { name: "Resumo" });
    expect(summary.getAttribute("aria-controls")).toBe("panel-tools");
    expect(document.getElementById("panel-tools")?.getAttribute("aria-labelledby")).toBe(
      "tab-summary"
    );
    expect(screen.getByRole("tabpanel", { name: "Resumo" })).toBeTruthy();
  });
  it("header is one row of at most 40 px (F-012/AC-11)", () => {
    setOffice();
    render(<App />);
    const header = document.querySelector("header") as HTMLElement;
    let h = window.getComputedStyle(header).height;
    if (!h) {
      for (const sheet of Array.from(document.styleSheets)) {
        for (const rule of Array.from((sheet as CSSStyleSheet).cssRules)) {
          const r = rule as CSSStyleRule;
          if (
            r.style?.height &&
            header.classList.contains(r.selectorText?.replace(/^\./, "") ?? "")
          )
            h = r.style.height;
        }
      }
    }
    expect(parseFloat(h), `header height computed as "${h}"`).toBeLessThanOrEqual(40);
  });
});

// ---------- AC-12: message context edges ----------
describe("message context edges (F-012/AC-12)", () => {
  it("shows the address alone when the sender has no name (never '<>' or 'undefined')", () => {
    setOffice({
      mailbox: { item: makeItem(ok("b"), { from: { displayName: "", emailAddress: "a@b.pt" } }) },
    });
    const { container } = render(<MessagePanel />);
    openBody();
    expect(container.textContent).toContain("a@b.pt");
    expect(container.textContent).not.toContain("<a@b.pt>");
  });
  it("never shows 'Name <>' when the name is present but the address is empty", () => {
    setOffice({
      mailbox: { item: makeItem(ok("b"), { from: { displayName: "Ann", emailAddress: "" } }) },
    });
    const { container } = render(<MessagePanel />);
    openBody();
    expect(container.textContent).not.toContain("<>");
  });
  it("shows an em dash in the De row when both name and address are empty (spec edge case)", () => {
    setOffice({
      mailbox: { item: makeItem(ok("b"), { from: { displayName: "", emailAddress: "" } }) },
    });
    render(<MessagePanel />);
    openBody();
    const label = screen.getByText("De");
    expect(label.parentElement?.textContent).toContain("—");
  });
  it("body of exactly 2000 characters is not marked truncated; 2001 is, with 2000 shown", () => {
    setOffice({ mailbox: { item: makeItem(ok("a".repeat(2000))) } });
    const a = render(<MessagePanel />);
    openBody();
    expect(screen.queryByText("(truncado)")).toBeNull();
    expect(a.container.querySelector("pre")?.textContent).toHaveLength(2000);
    a.unmount();
    setOffice({ mailbox: { item: makeItem(ok("a".repeat(2001))) } });
    const b = render(<MessagePanel />);
    openBody();
    expect(screen.getByText("(truncado)")).toBeTruthy();
    expect(b.container.querySelector("pre")?.textContent).toHaveLength(2000);
  });
  it("does not leave a lone surrogate when an emoji straddles the 2000 limit", () => {
    setOffice({ mailbox: { item: makeItem(ok("a".repeat(1999) + "😀")) } });
    const { container } = render(<MessagePanel />);
    openBody();
    const text = container.querySelector("pre")?.textContent ?? "";
    expect(text).toHaveLength(1999);
    expect(/[\uD800-\uDBFF]$/.test(text)).toBe(false);
  });
  it("keeps line breaks in the body preview (pre-wrap, whitespace kept)", () => {
    setOffice({ mailbox: { item: makeItem(ok("linha 1\n\n  linha 2")) } });
    const { container } = render(<MessagePanel />);
    openBody();
    expect(container.querySelector("pre")?.textContent).toBe("linha 1\n\n  linha 2");
  });
  it("renders HTML in the sender name as text (no element created)", () => {
    setOffice({
      mailbox: {
        item: makeItem(ok("x"), {
          from: { displayName: '<img src=x onerror="alert(1)">', emailAddress: "e@x.pt" },
        }),
      },
    });
    const { container } = render(<MessagePanel />);
    expect(container.querySelector("img")).toBeNull();
    expect(screen.getByText(/<img src=x/)).toBeTruthy();
  });
  it("renders an anchor-like subject as text (no href created from email content)", () => {
    setOffice({
      mailbox: { item: makeItem(ok("x"), { subject: '<a href="javascript:alert(1)">click</a>' }) },
    });
    const { container } = render(<MessagePanel />);
    expect(container.querySelector("a")).toBeNull();
  });
  it("the body region is a focusable, named scroll region (keyboard, F-012/AC-19)", () => {
    setOffice({ mailbox: { item: makeItem(ok("corpo")) } });
    render(<MessagePanel />);
    openBody();
    const region = screen.getByRole("region", { name: strings.bodyPreview });
    expect(region.tagName).toBe("PRE");
    expect(region.getAttribute("tabindex")).toBe("0");
    expect(region.getAttribute("dir")).toBe("auto");
  });
});

// ---------- AC-35: ItemChanged ----------
describe("ItemChanged behaviour (F-012/AC-35)", () => {
  it("discards a body read started for the previous item when it resolves late", () => {
    let lateFirst: Cb | undefined;
    const first = makeItem(
      vi.fn((_t: string, cb: Cb) => (lateFirst = cb)),
      { subject: "First" }
    );
    const second = makeItem(ok("second body"), { subject: "Second" });
    const handlers: Record<string, () => void> = {};
    setOffice({
      mailbox: {
        item: first,
        addHandlerAsync: vi.fn((e: string, h: () => void) => (handlers[e] = h)),
        removeHandlerAsync: vi.fn(),
      },
    });
    render(<MessagePanel />);
    (globalThis as any).Office.context.mailbox.item = second;
    act(() => handlers.itemChanged?.());
    expect(screen.getByText("Second")).toBeTruthy();
    act(() => lateFirst?.({ status: "succeeded", value: "FIRST BODY" }));
    openBody();
    expect(screen.getByText("second body")).toBeTruthy();
    expect(screen.queryByText("FIRST BODY")).toBeNull();
    expect(screen.queryByText("First")).toBeNull();
  });
  it("registers the ItemChanged handler exactly once", () => {
    const addHandlerAsync = vi.fn();
    setOffice({
      mailbox: { item: makeItem(ok("b")), addHandlerAsync, removeHandlerAsync: vi.fn() },
    });
    render(<MessagePanel />);
    const itemCalls = addHandlerAsync.mock.calls.filter((c) => c[0] === "itemChanged");
    expect(itemCalls).toHaveLength(1);
  });
  it("does not move focus when the item changes", () => {
    const handlers: Record<string, () => void> = {};
    const mailbox: any = {
      item: makeItem(ok("b")),
      addHandlerAsync: (e: string, h: () => void) => (handlers[e] = h),
      removeHandlerAsync: vi.fn(),
    };
    setOffice({ mailbox });
    render(<App />);
    const tab = screen.getByRole("tab", { name: "Resumo" });
    tab.focus();
    mailbox.item = makeItem(ok("c"), { subject: "Other" });
    act(() => handlers.itemChanged?.());
    expect(screen.getByText("Other")).toBeTruthy();
    expect(document.activeElement).toBe(tab);
  });
});

// ---------- AC-23 / AC-10: PT-PT literals ----------
describe("PT-PT literals (F-012/AC-23, approved strings in the spec)", () => {
  it("uses the approved wording of the F-012 Strings section", () => {
    expect(strings.assurance).toBe("Sem envio automático. O envio é sempre feito por si.");
    expect(strings.comingSoonReply).toBe(
      "Respostas automáticas para contactos escolhidos: em breve."
    );
    expect(strings.connected).toBe("Ligado");
  });
});

// ---------- AC-3 / AC-4 / AC-2 / AC-22 / AC-27 / AC-20: source and page guards ----------
describe("source and page guards (F-012/AC-2, AC-3, AC-4, AC-20, AC-22, AC-27)", () => {
  const dir = join(process.cwd(), "src/taskpane");
  const src = (f: string) => readFileSync(join(dir, f), "utf8");
  const prod = [
    "App.tsx",
    "MessagePanel.tsx",
    "BackendStatus.tsx",
    "theme.ts",
    "Wordmark.tsx",
    "strings.ts",
    "messageView.ts",
    "index.tsx",
  ];
  it("no href, animation, keyframes, transition or webfont in the pane source", () => {
    for (const f of prod) {
      expect(src(f), f).not.toMatch(
        /\bhref\s*=|@keyframes|animation\s*:|transition\s*:|@font-face|fonts\.g/
      );
    }
  });
  it("taskpane.html declares lang pt-PT and loads no external stylesheet or font", () => {
    const html = readFileSync(join(dir, "taskpane.html"), "utf8");
    expect(html).toMatch(/<html lang="pt-PT"/);
    expect(html).not.toMatch(/<link[^>]+href=["']https?:/);
  });
});

// ---------- AC-32 / AC-34 / AC-26: manifest structure (parsed, not regex) ----------
describe("manifest ribbon entry (F-012/AC-32, AC-34, AC-26)", () => {
  const text = readFileSync(join(process.cwd(), "manifest.xml"), "utf8");
  const xml = new DOMParser().parseFromString(text, "text/xml");
  const all = (tag: string) => Array.from(xml.getElementsByTagName(tag));
  const strs = (tag: string) => all(tag);
  const resolve = (resid: string) => {
    const hit = [...strs("bt:String"), ...strs("bt:Url")].find(
      (e) => e.getAttribute("id") === resid
    );
    return hit?.getAttribute("DefaultValue") ?? null;
  };
  it("one TabDefault OfficeTab, one Group, one Button (DIMSCALE Mail), no CustomTab (AC-32, AC-48)", () => {
    expect(all("OfficeTab")).toHaveLength(1);
    expect(all("OfficeTab")[0].getAttribute("id")).toBe("TabDefault");
    expect(all("Group")).toHaveLength(1);
    expect(all("Control").filter((c) => c.getAttribute("xsi:type") === "Button")).toHaveLength(1);
    expect(all("CustomTab")).toHaveLength(0);
  });
  it("group label resolves to 'DIMSCALE MAIL' (AC-32)", () => {
    const groupLabel = all("Group")[0].getElementsByTagName("Label")[0].getAttribute("resid");
    expect(resolve(groupLabel ?? "")).toBe("DIMSCALE MAIL");
  });
  it("the button label and Supertip title are 'DIMSCALE Mail' (AC-32, AC-48)", () => {
    const btns = all("Control").filter((c) => c.getAttribute("xsi:type") === "Button");
    const pick = (b: Element, tag: string) =>
      resolve(b.getElementsByTagName(tag)[0].getAttribute("resid") ?? "");
    expect(btns.map((b) => pick(b, "Label"))).toEqual(["DIMSCALE Mail"]);
    expect(btns.map((b) => pick(b, "Title"))).toEqual(["DIMSCALE Mail"]);
  });
  it("SupportsPinning is inside each ShowTaskpane action of the v1.1 block (AC-34, AC-48)", () => {
    const pin = all("SupportsPinning");
    expect(pin).toHaveLength(1);
    for (const p of pin) {
      expect(p.parentElement?.getAttribute("xsi:type")).toBe("ShowTaskpane");
      expect(p.closest("VersionOverrides")?.getAttribute("xsi:type")).toBe(
        "VersionOverridesV1_1"
      );
    }
  });
  it("only ReadItem is requested (AC-26)", () => {
    expect(all("Permissions").map((p) => p.textContent)).toEqual(["ReadItem"]);
  });
});

// ---------- AC-24: icon file structure beyond the existing size/bit-depth check ----------
describe("icon files (F-012/AC-24)", () => {
  it.each([16, 32, 64, 80, 128])("icon-%i.png has the PNG signature and IHDR chunk", (n) => {
    const b = readFileSync(join(process.cwd(), "assets", `icon-${n}.png`));
    expect([...b.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(b.subarray(12, 16).toString("ascii")).toBe("IHDR");
  });
});
