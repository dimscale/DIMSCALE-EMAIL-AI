/* global console, fetch */
import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { inspect } from "node:util";
import { join } from "node:path";

// QA (independent) for F-003 amendment (d), AC-54, AC-56, AC-57, AC-58 on the pane side.
// Mocked Office, auth and fetch only. Not a live Outlook run.
vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handler: (() => void) | undefined;
let fetchMock: ReturnType<typeof vi.fn>;
const E = (n: number) => "😀".repeat(n); // one code point, two UTF-16 units each
const CONTRACT = ["action", "body", "draft", "fromDomain", "fromName", "instruction", "language", "subject", "tone"];

const makeItem = (subject: string) => ({
  subject,
  from: { displayName: "Ana", emailAddress: "ana@Exemplo.pt" },
  dateTimeCreated: new Date(2026, 0, 2),
  body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: "corpo" }) },
  displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
  displayReplyForm: vi.fn(),
});
const setup = () => {
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      mailbox: {
        item: makeItem("Assunto"),
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handler = h;
        }),
        removeHandlerAsync: vi.fn(),
      },
    },
  };
};
const current = () => (globalThis as any).Office.context.mailbox.item;
const res = (reply: string) => ({ ok: true, status: 200, json: async () => ({ ok: true, action: "reply", reply }) });
const fail = (status: number, error: string) => ({ ok: false, status, json: async () => ({ ok: false, error }) });
const btn = (name: string) => screen.getByRole("button", { name }) as HTMLButtonElement;
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(btn(name));
  });
};
const box = () => screen.getByPlaceholderText("Ex.: mais curto, pede o prazo") as HTMLInputElement;
const type = (v: string) => fireEvent.change(box(), { target: { value: v } });
const enter = async () => {
  await act(async () => {
    fireEvent.keyDown(box(), { key: "Enter" });
  });
};
const bodyOf = (i: number) => JSON.parse(fetchMock.mock.calls[i][1].body);
const logSpies = () => (["log", "info", "warn", "error", "debug", "trace"] as const).map((m) => vi.spyOn(console, m).mockImplementation(() => {}));

beforeEach(async () => {
  fetchMock = vi.fn().mockResolvedValue(res("v0"));
  vi.stubGlobal("fetch", fetchMock);
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "TOK" });
  setup();
  render(<App />);
  fireEvent.click(screen.getByRole("tab", { name: "Resposta" }));
  await click("Sugerir resposta");
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("AC-54/AC-58 (QA): request body contract", () => {
  it("a revision sends exactly the nine contract keys, trimmed instruction, shown draft; a default request sends seven", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("  mais curto ");
    await click("Enviar alterações");
    expect(Object.keys(bodyOf(1)).sort()).toEqual(CONTRACT);
    expect(Object.keys(bodyOf(0)).sort()).toEqual(CONTRACT.filter((k) => k !== "draft" && k !== "instruction"));
    expect(bodyOf(1).instruction).toBe("mais curto");
    expect(bodyOf(1).draft).toBe("v0");
    expect(bodyOf(1).action).toBe("reply");
    expect(bodyOf(1).subject).toBe("Assunto");
  });
});

describe("whitespace-only instruction (AC-54 QA)", () => {
  it("blank, tab, NBSP and BOM-only instructions: button disabled and Enter sends nothing", async () => {
    for (const v of ["", " ", "\t \n", "  ", "﻿"]) {
      type(v);
      expect(btn("Enviar alterações").disabled).toBe(true);
      await enter();
    }
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe("Unicode at the limits, UTF-16 units (AC-53 QA, pane)", () => {
  it("a 2,000-unit draft (1,000 emoji) is shown and sent whole; a 300-unit instruction (150 emoji) is sent", async () => {
    fetchMock.mockResolvedValue(res(E(1000)));
    await click("Sugerir resposta");
    expect(box()).toBeTruthy(); // 2,000 units is allowed
    fetchMock.mockResolvedValue(res("ok"));
    type(E(150));
    await click("Enviar alterações");
    expect(bodyOf(2).draft.length).toBe(2000);
    expect(bodyOf(2).draft).toBe(E(1000));
    expect(bodyOf(2).instruction.length).toBe(300);
  });

  it("a 1,001-emoji suggestion (2,002 units, only 1,001 code points) disables the box", async () => {
    fetchMock.mockResolvedValue(res(E(1001)));
    await click("Sugerir resposta");
    expect(screen.queryByPlaceholderText("Ex.: mais curto, pede o prazo")).toBeNull();
    expect(screen.getByText(/demasiado longa/)).toBeTruthy();
  });
});

describe("history and the reply form (AC-56 QA)", () => {
  it("a new item starts with no suggestion, no Versão anterior and an empty box; its first suggestion has no revision fields", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("mais curto");
    await click("Enviar alterações");
    expect(btn("Versão anterior")).toBeTruthy();
    (globalThis as any).Office.context.mailbox.item = makeItem("Outro");
    act(() => handler?.());
    await waitFor(() => expect(screen.queryByText("v1")).toBeNull());
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull();
    expect(screen.queryByPlaceholderText("Ex.: mais curto, pede o prazo")).toBeNull(); // no suggestion, no refine box
    fetchMock.mockResolvedValue(res("novo"));
    await click("Sugerir resposta");
    expect(screen.getByText("novo")).toBeTruthy();
    expect(box().value).toBe("");
    expect(bodyOf(fetchMock.mock.calls.length - 1)).not.toHaveProperty("draft");
  });

  it("Abrir resposta para revisão sends the shown version: the new one, then the restored one", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("mais curto");
    await click("Enviar alterações");
    await click("Abrir resposta para revisão");
    const item = current();
    expect(item.displayReplyFormAsync.mock.calls[0][0]).toBe("v1");
    await click("Versão anterior");
    await click("Abrir resposta para revisão");
    expect(item.displayReplyFormAsync.mock.calls[1][0]).toBe("v0");
    expect(item.displayReplyFormAsync).toHaveBeenCalledTimes(2);
  });

  it("after Versão anterior, the next revision sends the restored text as draft", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("a");
    await click("Enviar alterações");
    await click("Versão anterior");
    fetchMock.mockResolvedValue(res("v2"));
    type("b");
    await click("Enviar alterações");
    expect(bodyOf(2).draft).toBe("v0");
    expect(bodyOf(2).instruction).toBe("b");
    expect(screen.getByText("v2")).toBeTruthy();
  });
});

describe("failed revision keeps the shown text and the typed instruction (AC-54 QA)", () => {
  it("a 502 keeps v0 and the instruction, shows one alert, offers no retry and no history entry", async () => {
    fetchMock.mockResolvedValue(fail(502, "ai_output_invalid"));
    type("mais curto");
    await click("Enviar alterações");
    expect(screen.getByRole("alert")).toBeTruthy();
    expect(screen.getByText("v0")).toBeTruthy();
    expect(box().value).toBe("mais curto");
    expect(screen.queryByRole("button", { name: "Tentar novamente" })).toBeNull();
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull();
  });
});

describe("AC-57 (QA): nothing is logged on the revision flow (pane)", () => {
  it("success, 422, network failure and item change during a revision: no console call, no instruction or draft in any argument", async () => {
    const ss = logSpies();
    fetchMock.mockResolvedValue(res("MARK_DRAFT_v0"));
    await click("Sugerir resposta");
    fetchMock.mockResolvedValue(res("MARK_REPLY_v1"));
    type("MARK_INSTR_1");
    await click("Enviar alterações");
    fetchMock.mockResolvedValue(fail(422, "payload_invalid"));
    type("MARK_INSTR_2");
    await click("Enviar alterações");
    fetchMock.mockRejectedValueOnce(new Error("MARK_NET_ERR"));
    await click("Enviar alterações");
    fetchMock.mockImplementation(
      (_u: string, init: { signal: AbortSignal }) =>
        new Promise((_r, rej) => init.signal.addEventListener("abort", () => rej(new Error("MARK_ABORT"))))
    );
    type("MARK_INSTR_3");
    await click("Enviar alterações");
    (globalThis as any).Office.context.mailbox.item = makeItem("Outro");
    act(() => handler?.());
    await waitFor(() => expect(screen.queryByText("MARK_REPLY_v1")).toBeNull());
    const sent = fetchMock.mock.calls.map((c) => String(c[1].body ?? "")).join("\n");
    expect(sent).toContain("MARK_INSTR_1"); // the markers did reach the request, so the check below is meaningful
    const calls = ss.flatMap((s) => s.mock.calls);
    expect(calls).toHaveLength(0);
    expect(inspect(calls, { depth: 6 })).not.toMatch(/MARK_/);
  });
});

describe("AC-57 (QA): source guard on the revision code path", () => {
  it("AiTools.tsx, analyzeClient.ts, replyForm.ts, messageView.ts and strings.ts contain no console call", () => {
    for (const f of ["AiTools.tsx", "analyzeClient.ts", "replyForm.ts", "messageView.ts", "strings.ts"]) {
      const src = readFileSync(join(process.cwd(), "src", "taskpane", f), "utf8");
      expect(src, f).not.toMatch(/console\./);
    }
  });
});
