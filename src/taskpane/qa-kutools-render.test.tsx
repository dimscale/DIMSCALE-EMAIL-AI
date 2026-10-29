import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

// QA (independent): AI text renders as text only (F-003/AC-20, AC-41; F-012/AC-3). Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const HOSTILE_POINT = '<img src=x onerror="alert(1)"><script>alert(2)</script>';
const HOSTILE_LINK = "[clique](javascript:alert(3)) http://evil.example";
const RTL = "‮evil RTL override‬ مرحبا بالعالم 📧";
const LONG = "x".repeat(3000); // one unbroken token
const SUMMARY = {
  points: [HOSTILE_POINT, HOSTILE_LINK, LONG],
  type: "request",
  priority: "high",
  tone: "negative",
  deadline: '<a href="http://evil.example">sexta</a>',
  requestedAction: "javascript:alert(4) " + RTL,
};

const makeItem = (body: string, o: object = {}) => ({
  subject: "Assunto",
  from: { displayName: "Ana", emailAddress: "ana@exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }) },
  displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
  displayReplyForm: vi.fn(),
  ...o,
});
const setup = (item: unknown) => {
  office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      roamingSettings: { get: vi.fn(() => undefined), set: vi.fn(), saveAsync: vi.fn() },
      mailbox: { item, addHandlerAsync: vi.fn(), removeHandlerAsync: vi.fn() },
    },
  };
  (globalThis as any).Office = office;
};
const json = (status: number, body: unknown) => ({
  ok: status < 300,
  status,
  json: async () => body,
});
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(screen.getByRole("button", { name }));
  });
};
const tab = (name: string) => fireEvent.click(screen.getByRole("tab", { name }));
// Elements that would mean HTML was interpreted. Fluent icons are <svg>, so svg is not listed.
const RISKY = "img, script, iframe, object, embed, a[href], a, link, meta, style, input, form";

beforeEach(() => {
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("summary renders as text (F-003/AC-20, F-012/AC-3)", () => {
  it("hostile points, deadline and requested action are text nodes; no element is created from them; dir=auto", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    const { container } = render(<App />);
    await click("Resumir");
    expect(container.querySelectorAll(RISKY)).toHaveLength(0);
    expect(screen.getByText(HOSTILE_POINT).getAttribute("dir")).toBe("auto");
    expect(screen.getByText(HOSTILE_LINK).tagName).toBe("LI");
    expect(screen.getByText(LONG).tagName).toBe("LI");
    expect(screen.getByText('Prazo: <a href="http://evil.example">sexta</a>').tagName).toBe("P");
    expect(screen.getByText("Pedido: javascript:alert(4) " + RTL)).toBeTruthy();
    expect(container.querySelectorAll("li")).toHaveLength(3);
  });

  it("long unbroken text gets overflow-wrap: anywhere (computed style of the bullet)", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    render(<App />);
    await click("Resumir");
    const li = screen.getByText(LONG);
    expect(window.getComputedStyle(li).overflowWrap).toBe("anywhere");
    expect(window.getComputedStyle(li.parentElement as Element).overflowWrap).toBe("anywhere");
  });

  it("unknown or prototype-like enum values give no badge (own-property lookup)", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(
      json(200, {
        ok: true,
        action: "summarize",
        summary: {
          ...SUMMARY,
          type: "constructor",
          priority: "__proto__",
          tone: "toString",
          deadline: null,
          requestedAction: null,
          points: ["a", "b", "c"],
        },
      })
    );
    const { container } = render(<App />);
    await click("Resumir");
    const text = container.textContent ?? "";
    expect(text).not.toMatch(/\[object|function |Object\]/);
    expect(text).not.toMatch(/Pedido|Prioridade|Tom /);
    expect(text).not.toContain("Prazo:");
    expect(text).not.toContain("Pedido:");
  });

  it("copy writes the hostile bullets literally (no HTML entity or tag conversion)", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "summarize", summary: SUMMARY }));
    const writeText = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    await click("Resumir");
    await click("Copiar");
    expect(writeText).toHaveBeenCalledWith(`• ${HOSTILE_POINT}\n• ${HOSTILE_LINK}\n• ${LONG}`);
  });
});

describe("reply renders as text (F-003/AC-20, AC-24)", () => {
  it("HTML, RTL and a 3,000-character unbroken reply render literally, pre-wrap, dir=auto", async () => {
    setup(makeItem("corpo"));
    const reply = `<svg onload="alert(5)"><b>x</b></svg> & ' " مرحبا\n` + "y".repeat(3000);
    fetchMock.mockResolvedValue(json(200, { ok: true, action: "reply", reply }));
    const { container } = render(<App />);
    tab("Resposta");
    await click("Sugerir resposta");
    expect(container.querySelectorAll(RISKY)).toHaveLength(0);
    expect(container.querySelector("svg[onload], b")).toBeNull();
    const p = screen.getByText(/onload="alert\(5\)"/);
    expect(p.textContent).toBe(reply); // the whole text, newline included, nothing dropped or converted
    expect(p.getAttribute("dir")).toBe("auto");
    expect(window.getComputedStyle(p).whiteSpace).toBe("pre-wrap");
    expect(window.getComputedStyle(p).overflowWrap).toBe("anywhere");
  });
});

describe("translation renders as text (F-003/AC-41)", () => {
  it("hostile subject and body render literally; the subject line only when non-empty; no link element", async () => {
    setup(makeItem("corpo"));
    const subject = "Pedido <a href='http://evil.example'>x</a>";
    const body = `<img src=x onerror=alert(6)>\n${RTL}\n[link](http://evil.example)`;
    fetchMock.mockResolvedValue(
      json(200, { ok: true, action: "translate", translation: { subject, body } })
    );
    const { container } = render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(container.querySelectorAll(RISKY)).toHaveLength(0);
    const subjectLine = screen.getByText(`Assunto: ${subject}`);
    expect(subjectLine.getAttribute("dir")).toBe("auto");
    const pre = screen.getByText((_, el) => el?.tagName === "P" && (el.textContent ?? "") === body);
    expect(pre.getAttribute("dir")).toBe("auto");
  });

  it("a whitespace-only subject shows no subject line", async () => {
    setup(makeItem("corpo"));
    fetchMock.mockResolvedValue(
      json(200, {
        ok: true,
        action: "translate",
        translation: { subject: "   ", body: "so corpo" },
      })
    );
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    expect(screen.queryByText(/^Assunto:/)).toBeNull();
    expect(screen.getByText("so corpo")).toBeTruthy();
  });
});

describe("message context renders hostile subject and sender as text (F-012/AC-3)", () => {
  it("a subject and display name with markup and RTL are text; no element from them", async () => {
    setup(
      makeItem("corpo", {
        subject: `<img src=x onerror=alert(7))> ${RTL}`,
        from: { displayName: "<b>Ana</b>", emailAddress: "ana@exemplo.pt" },
      })
    );
    const { container } = render(<App />);
    await act(async () => undefined);
    expect(container.querySelector("img, b")).toBeNull();
    expect(screen.getByText(`<img src=x onerror=alert(7))> ${RTL}`)).toBeTruthy();
  });
});
