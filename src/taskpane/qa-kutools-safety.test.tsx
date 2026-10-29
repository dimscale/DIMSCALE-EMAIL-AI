/* global console */
import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, relative, sep } from "node:path";

// QA (independent): safety invariants (CLAUDE.md rules 1, 3, 5; F-003/AC-12, AC-25; F-012/AC-3, AC-4). Mocks only.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string; error?: unknown }) => void;

// Sentinels that must never reach a log line: body, subject, name, address, domain, token, error text.
const SENTINELS = [
  "SEGREDO-CORPO-123",
  "ASSUNTO-SECRETO-456",
  "ana@exemplo.pt",
  "Ana Silva",
  "exemplo.pt",
  "TOK-SECRET-XYZ",
  "SEGREDO-TOKEN-789",
  "SEGREDO-ERRO-000",
];
const BODY = `Olá ${SENTINELS[0]} de ${SENTINELS[3]} <${SENTINELS[2]}>`;
const METHODS = ["log", "info", "warn", "error", "debug", "trace"] as const;

let spies: ReturnType<typeof vi.spyOn>[] = [];
let office: any;
let fetchMock: ReturnType<typeof vi.fn>;

const makeItem = (body: string, o: object = {}) => ({
  subject: SENTINELS[1],
  from: { displayName: SENTINELS[3], emailAddress: SENTINELS[2] },
  dateTimeCreated: new Date(2026, 0, 2),
  body: {
    getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: body }),
  },
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

beforeEach(() => {
  spies = METHODS.map((m) => vi.spyOn(console, m).mockImplementation(() => undefined));
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK-SECRET-XYZ" });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  spies.forEach((s) => s.mockRestore());
});

const logged = (): string[] =>
  spies.reduce<string[]>(
    (acc, s) =>
      acc.concat(
        s.mock.calls.map((args: unknown[]) =>
          args
            .map((a) => {
              if (a instanceof Error) return `${a.name} ${a.message} ${a.stack ?? ""}`;
              try {
                return typeof a === "string" ? a : (JSON.stringify(a) ?? String(a));
              } catch {
                return String(a);
              }
            })
            .join(" ")
        )
      ),
    []
  );

describe("no console output carries email content, tokens or error text (F-003/AC-12)", () => {
  it("canary: the spy would catch a logged sentinel (so the checks below can fail)", () => {
    console.error("canary", SENTINELS[0]);
    expect(logged().some((l) => l.includes(SENTINELS[0]))).toBe(true);
    spies.forEach((sp) => sp.mockClear());
  });

  it("success, server error, network failure, clipboard failure, reply-form failure, malformed JSON and body-read failure: console stays free of content", async () => {
    // 1. body read fails with an error that carries the body text
    setup(
      makeItem(BODY, {
        body: {
          getAsync: (_t: string, cb: Cb) =>
            cb({ status: "failed", error: new Error(SENTINELS[0]) }),
        },
      })
    );
    const first = render(<App />);
    expect(screen.getByRole("alert")).toBeTruthy();
    first.unmount();

    // 2. summary succeeds
    setup(makeItem(BODY));
    fetchMock.mockResolvedValueOnce(
      json(200, {
        ok: true,
        action: "summarize",
        summary: {
          points: ["a", "b", "c"],
          type: "request",
          priority: "high",
          tone: "neutral",
          deadline: null,
          requestedAction: null,
        },
      })
    );
    const writeText = vi.fn().mockRejectedValue(new Error(SENTINELS[1]));
    vi.stubGlobal("navigator", { clipboard: { writeText } });
    render(<App />);
    await click("Resumir");
    await click("Copiar");

    // 3. reply: network failure whose message carries the body
    tab("Resposta");
    fetchMock.mockRejectedValueOnce(new TypeError(`Failed to fetch ${SENTINELS[0]}`));
    await click("Sugerir resposta");
    // 4. reply succeeds, reply form throws with content in the message
    fetchMock.mockResolvedValueOnce(
      json(200, { ok: true, action: "reply", reply: `Olá ${SENTINELS[3]}` })
    );
    await click("Sugerir resposta");
    const item = office.context.mailbox.item as any;
    item.displayReplyFormAsync = vi.fn(() => {
      throw new Error(SENTINELS[0]);
    });
    await click("Abrir resposta para revisão");

    // 5. translate: malformed JSON and a 500 whose body carries an error text
    tab("Traduzir");
    fetchMock.mockResolvedValueOnce({
      ok: false,
      status: 502,
      json: async () => {
        throw new SyntaxError(SENTINELS[7]);
      },
    });
    await click("Traduzir");
    fetchMock.mockResolvedValueOnce(
      json(500, { ok: false, error: "server_error", detail: SENTINELS[7] })
    );
    await click("Traduzir");

    // 6. no sign-in: token error carries a sentinel code
    (getTokenSilent as any).mockResolvedValue({ kind: "error", code: SENTINELS[6] });
    await click("Traduzir");

    expect(screen.queryByText(/SEGREDO-ERRO-000/)).toBeNull(); // the raw detail is never rendered either
    const lines = logged();
    for (const s of SENTINELS) {
      for (const line of lines) expect(line.includes(s), `console contained "${s}"`).toBe(false);
    }
  });

  it("a successful run with every card used logs nothing that matches the content (count is not asserted, content is)", async () => {
    setup(makeItem(BODY));
    fetchMock.mockResolvedValue(
      json(200, {
        ok: true,
        action: "translate",
        translation: { subject: SENTINELS[1], body: BODY },
      })
    );
    render(<App />);
    tab("Traduzir");
    await click("Traduzir");
    for (const line of logged()) for (const s of SENTINELS) expect(line.includes(s)).toBe(false);
  });
});

// Static walk of every non-test source file the pane ships (prototype and node_modules excluded).
const ROOT = join(process.cwd(), "src", "taskpane");
const walk = (dir: string): string[] =>
  readdirSync(dir).reduce<string[]>((acc, name) => {
    const p = join(dir, name);
    if (name === "node_modules" || name === "prototype") return acc;
    if (statSync(p).isDirectory()) return acc.concat(walk(p));
    if (!/\.(ts|tsx|html)$/.test(name) || /\.test\.(ts|tsx)$/.test(name)) return acc;
    return acc.concat([p]);
  }, []);
const files = walk(ROOT);
const text = (f: string) => readFileSync(f, "utf8");
const name = (f: string) => relative(ROOT, f).split(sep).join("/");

describe("static guards over every pane source file (F-003/AC-25, F-012/AC-3, AC-4)", () => {
  it("the walk covers the expected AI files", () => {
    for (const f of [
      "AiTools.tsx",
      "ToolCard.tsx",
      "AutoSummarySwitch.tsx",
      "autoSummary.ts",
      "prefs.ts",
      "analyzeClient.ts",
      "messageView.ts",
      "App.tsx",
      "MessagePanel.tsx",
      "index.tsx",
      "config.ts",
      "strings.ts",
      "replyForm.ts",
      "taskpane.html",
    ]) {
      expect(files.map(name)).toContain(f);
    }
  });

  it.each([
    ["console", /\bconsole\./],
    ["web storage", /\b(localStorage|sessionStorage)\b/],
    ["cookies / IndexedDB", /document\.cookie|indexedDB/],
    [
      "HTML injection",
      /dangerouslySetInnerHTML|innerHTML|insertAdjacentHTML|outerHTML|document\.write/,
    ],
    ["eval", /\beval\(|new Function\(/],
    ["location.search / location.hash", /location\.(search|hash)/],
    [
      "send / reply-all / new-message APIs",
      /Mail\.Send|displayReplyAll|displayNewMessage|sendAsync/,
    ],
    ["window.open / location.href assignment", /window\.open|location\.href\s*=/],
  ])("no %s anywhere in the pane source", (_label, pattern) => {
    const hits = files.filter((f) => pattern.test(text(f))).map(name);
    expect(hits).toEqual([]);
  });

  it("no href or userProfile in the AI files (F-012 AC-3: no link from email content; the identity is not sent)", () => {
    const aiFiles = [
      "AiTools.tsx",
      "ToolCard.tsx",
      "AutoSummarySwitch.tsx",
      "autoSummary.ts",
      "prefs.ts",
      "analyzeClient.ts",
      "messageView.ts",
      "replyForm.ts",
      "MessagePanel.tsx",
    ];
    for (const f of files.filter((x) => aiFiles.includes(name(x)))) {
      expect(text(f), name(f)).not.toMatch(/\bhref\s*=|userProfile/);
    }
  });

  it("itemId is read only as the in-memory dedupe key in AiTools, never in the request module", () => {
    expect(text(join(ROOT, "analyzeClient.ts"))).not.toMatch(/itemId/);
    const aiTools = text(join(ROOT, "AiTools.tsx"));
    const uses = aiTools.split("\n").filter((l) => /itemId/.test(l));
    expect(uses).toHaveLength(1);
    expect(uses[0]).toMatch(/const id = view\.itemId/);
  });

  it("the only AI network call is analyzeClient.ts (no fetch in the AI cards)", () => {
    for (const f of [
      "AiTools.tsx",
      "ToolCard.tsx",
      "AutoSummarySwitch.tsx",
      "autoSummary.ts",
      "prefs.ts",
    ]) {
      expect(text(join(ROOT, f)), f).not.toMatch(/\bfetch\(/);
    }
    expect(text(join(ROOT, "analyzeClient.ts"))).toMatch(/fetchFn\(ANALYZE_URL/);
  });

  it("the entry decision reads only location.pathname (index.tsx)", () => {
    const idx = text(join(ROOT, "index.tsx"));
    expect(idx.match(/location\.[a-z]+/g)).toEqual(["location.pathname"]);
  });

  it("the resumo entry page is generated from the same template as taskpane.html (no inline script in the template)", () => {
    const wp = readFileSync(join(process.cwd(), "webpack.config.js"), "utf8");
    expect(wp).toMatch(/filename: "resumo\.html",\s*template: "\.\/src\/taskpane\/taskpane\.html"/);
    const tpl = text(join(ROOT, "taskpane.html"));
    expect(tpl).not.toMatch(/<script(?![^>]*\bsrc=)/);
  });

  it("the CopyWebpackPlugin rule for manifests copies only manifest.xml, never the test manifest (F-012 AC-37, AC-49)", () => {
    const wp = readFileSync(join(process.cwd(), "webpack.config.js"), "utf8");
    expect(wp).not.toMatch(/from:\s*"manifest\*\.xml"/);
    expect(wp).toMatch(/from:\s*"manifest\.xml"/);
  });
});
