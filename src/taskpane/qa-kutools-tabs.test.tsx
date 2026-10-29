import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): four tabs, ARIA wiring and focus (F-012/AC-19, AC-43). jsdom only: arrow keys are NOT VERIFIED here.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handlers: Set<() => void>;
let office: any;

const makeItem = (body: string) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
});
const setup = (item: unknown) => {
  handlers = new Set();
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: { get: vi.fn(() => undefined), set: vi.fn(), saveAsync: vi.fn() },
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handlers.add(h);
        }),
        removeHandlerAsync: vi.fn(),
      },
    },
  };
  (globalThis as any).Office = office;
};

beforeEach(() => {
  vi.stubGlobal("fetch", vi.fn());
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("tab list semantics (F-012/AC-43)", () => {
  it("one tablist with exactly four tabs in the order Resumo, Resposta, Traduzir, Automático, all native buttons", () => {
    setup(makeItem("corpo"));
    render(<App />);
    const list = screen.getByRole("tablist");
    const tabs = screen.getAllByRole("tab");
    expect(list.contains(tabs[0])).toBe(true);
    expect(tabs.map((t) => t.getAttribute("value"))).toEqual([
      "summary",
      "reply",
      "translate",
      "auto",
    ]);
    for (const n of ["Resumo", "Resposta", "Traduzir", "Automático"])
      expect(screen.getByRole("tab", { name: n })).toBeTruthy();
    for (const t of tabs) expect(t.tagName).toBe("BUTTON"); // Enter and Space activate native buttons in a browser
  });

  it("aria-selected is true on exactly one tab, and the tab's aria-controls points at an existing tabpanel", () => {
    setup(makeItem("corpo"));
    render(<App />);
    const selectedAt = () =>
      screen
        .getAllByRole("tab")
        .filter((t) => t.getAttribute("aria-selected") === "true")
        .map((t) => t.textContent);
    expect(selectedAt()).toEqual(["Resumo"]);
    for (const name of ["Resposta", "Traduzir", "Automático"]) {
      fireEvent.click(screen.getByRole("tab", { name }));
      expect(selectedAt()).toEqual([name]);
    }
    for (const t of screen.getAllByRole("tab")) {
      const target = t.getAttribute("aria-controls");
      expect(target).toBeTruthy();
      expect(document.getElementById(target as string)).not.toBeNull();
    }
  });

  it("each tab names the tabpanel it controls (aria-labelledby), and the Automático panel is its own panel", () => {
    setup(makeItem("corpo"));
    render(<App />);
    fireEvent.click(screen.getByRole("tab", { name: "Traduzir" }));
    expect(screen.getByRole("tabpanel", { name: "Traduzir" })).toBeTruthy();
    fireEvent.click(screen.getByRole("tab", { name: "Automático" }));
    const auto = document.getElementById("panel-auto") as HTMLElement;
    expect(auto.getAttribute("aria-labelledby")).toBe("tab-auto");
    expect(auto.hasAttribute("hidden")).toBe(false);
    expect(screen.getByRole("tabpanel", { name: "Automático" })).toBeTruthy();
  });

  it("the removed labels never render: Analisar, Responder, Escrever, the old heading and the old 'em breve' card", () => {
    setup(makeItem("corpo"));
    const { container } = render(<App />);
    for (const n of [
      "Analisar",
      "Responder",
      "Escrever",
      "DIMSCALE AI Assistant",
      "Melhorar texto e tom",
      "Resposta rápida",
    ]) {
      expect(container.textContent ?? "").not.toContain(n);
    }
  });

  it("the tab row is labelled (aria-label) in PT-PT", () => {
    setup(makeItem("corpo"));
    render(<App />);
    expect(screen.getByRole("tablist").getAttribute("aria-label")).toBe("Vistas");
  });
});

describe("focus (F-012/AC-19, AC-35)", () => {
  it("focus on a tab stays on that tab when the item changes (the tab row is not remounted)", async () => {
    setup(makeItem("A"));
    render(<App />);
    const tab = screen.getByRole("tab", { name: "Resposta" });
    tab.focus();
    office.context.mailbox.item = makeItem("B");
    await act(async () => handlers.forEach((h) => h()));
    expect(document.activeElement).toBe(tab);
  });

  it("focus on a card button is lost when the item changes (the card is remounted per item): recorded as an observation", async () => {
    setup(makeItem("A"));
    render(<App />);
    const btn = screen.getByRole("button", { name: "Resumir" });
    btn.focus();
    expect(document.activeElement).toBe(btn);
    office.context.mailbox.item = makeItem("B");
    await act(async () => handlers.forEach((h) => h()));
    // observed: focus falls back to the body; nothing is focused on purpose. Keyboard users lose their place.
    expect(document.activeElement).not.toBe(btn);
    expect(screen.getByRole("button", { name: "Resumir" })).toBeTruthy();
  });
});

describe("keyboard (NOT VERIFIED WITH MOCKS)", () => {
  it.todo(
    "arrow keys move between the four tabs (Fluent TabList, browser keyboard focus: NOT VERIFIED in jsdom)"
  );
  it.todo(
    "Enter and Space activate a focused tab in a browser (NOT VERIFIED: jsdom does not synthesise the native click)"
  );
});
