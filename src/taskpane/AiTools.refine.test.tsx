import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./BackendStatus", () => ({ default: () => <div role="status">stub</div> }));
vi.mock("./auth", () => ({ getTokenSilent: vi.fn(), isNaaSupported: () => true, signIn: vi.fn() }));
import App from "./App";
import { getTokenSilent } from "./auth";

type Cb = (r: { status: string; value?: string }) => void;
let handler: (() => void) | undefined;
let fetchMock: ReturnType<typeof vi.fn>;

const setup = () => {
  const item = {
    subject: "Assunto",
    from: { displayName: "Ana", emailAddress: "ana@Exemplo.pt" },
    dateTimeCreated: new Date(2026, 0, 2),
    body: { getAsync: (_t: string, cb: Cb) => cb({ status: "succeeded", value: "corpo" }) },
    displayReplyFormAsync: vi.fn((_h: string, cb: Cb) => cb({ status: "succeeded" })),
    displayReplyForm: vi.fn(),
  };
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: { isSetSupported: vi.fn(() => true) },
      mailbox: {
        item,
        addHandlerAsync: vi.fn((e: string, h: () => void) => {
          if (e === "itemChanged") handler = h;
        }),
        removeHandlerAsync: vi.fn(),
      },
    },
  };
};
const res = (reply: string) => ({
  ok: true,
  status: 200,
  json: async () => ({ ok: true, action: "reply", reply }),
});
const btn = (name: string) => screen.getByRole("button", { name }) as HTMLButtonElement;
const click = async (name: string) => {
  await act(async () => {
    fireEvent.click(btn(name));
  });
};
const box = () => screen.getByPlaceholderText("Ex.: mais curto, pede o prazo") as HTMLInputElement;
const type = (v: string) => fireEvent.change(box(), { target: { value: v } });
const bodyOf = (i: number) => JSON.parse(fetchMock.mock.calls[i][1].body);

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
});

describe("reply refinement (F-003 amendment (d), AC-54, AC-56)", () => {
  it("no request on render or typing; empty/blank disables the button", () => {
    expect(fetchMock).toHaveBeenCalledTimes(1); // only the Sugerir resposta click
    expect(btn("Enviar alterações").disabled).toBe(true);
    type("   ");
    expect(btn("Enviar alterações").disabled).toBe(true);
    type("mais curto");
    expect(btn("Enviar alterações").disabled).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(box().maxLength).toBe(300);
  });

  it("one click = one request with instruction + draft; replaces text and clears the box", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("  mais curto ");
    await click("Enviar alterações");
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(bodyOf(1)).toMatchObject({
      action: "reply",
      instruction: "mais curto",
      draft: "v0",
      language: "match",
      tone: "none",
      subject: "Assunto",
    });
    expect(bodyOf(0).instruction).toBeUndefined(); // AC-58
    expect(screen.getByText("v1")).toBeTruthy();
    expect(box().value).toBe("");
  });

  it("Enter sends exactly one request; disabled while pending", async () => {
    let resolve!: (v: unknown) => void;
    fetchMock.mockReturnValue(new Promise((r) => (resolve = r)));
    type("tira o agradecimento");
    await act(async () => {
      fireEvent.keyDown(box(), { key: "Enter" });
      fireEvent.keyDown(box(), { key: "Enter" });
    });
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(btn("Enviar alterações").disabled).toBe(true);
    expect(box().readOnly).toBe(true);
    expect(box().disabled).toBe(false);
    await act(async () => resolve(res("v1")));
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("item change aborts the in-flight refinement", async () => {
    fetchMock.mockImplementation(
      (_u: string, init: { signal: AbortSignal }) =>
        new Promise((_r, rej) =>
          init.signal.addEventListener("abort", () => rej(new Error("abort")))
        )
    );
    type("mais curto");
    await click("Enviar alterações");
    const signal = fetchMock.mock.calls[1][1].signal as AbortSignal;
    expect(signal.aborted).toBe(false);
    cleanup(); // pane unmounts the tools on ItemChanged
    expect(signal.aborted).toBe(true);
  });

  it("Versão anterior restores with zero requests; history capped at 5", async () => {
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull();
    for (let i = 1; i <= 7; i++) {
      fetchMock.mockResolvedValue(res(`v${i}`));
      type(`a${i}`);
      await click("Enviar alterações");
    }
    expect(fetchMock).toHaveBeenCalledTimes(8);
    expect(bodyOf(7).draft).toBe("v6");
    for (let i = 6; i >= 2; i--) {
      await click("Versão anterior");
      expect(screen.getByText(`v${i}`)).toBeTruthy();
    }
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull(); // v0, v1 dropped
    expect(fetchMock).toHaveBeenCalledTimes(8);
  });

  it("a new Sugerir resposta resets the history", async () => {
    type("x");
    await click("Enviar alterações");
    expect(btn("Versão anterior")).toBeTruthy();
    fetchMock.mockResolvedValue(res("novo"));
    await click("Sugerir resposta");
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull();
  });

  it("renders markup in the reply as text", async () => {
    fetchMock.mockResolvedValue(res("<img src=x onerror=alert(1)> <b>olá</b>"));
    type("x");
    await click("Enviar alterações");
    expect(screen.getByText("<img src=x onerror=alert(1)> <b>olá</b>")).toBeTruthy();
    expect(document.querySelector("img")).toBeNull();
  });

  it("focus stays on the input after the response, a status is announced, no extra request", async () => {
    fetchMock.mockResolvedValue(res("v1"));
    type("mais curto");
    box().focus();
    await act(async () => {
      fireEvent.keyDown(box(), { key: "Enter" });
    });
    expect(document.activeElement).toBe(box());
    expect(screen.getByText("Resposta atualizada").getAttribute("role")).toBe("status");
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("Versão anterior on the last step moves focus to the input", async () => {
    type("x");
    await click("Enviar alterações");
    btn("Versão anterior").focus();
    await click("Versão anterior");
    expect(screen.queryByRole("button", { name: "Versão anterior" })).toBeNull();
    expect(document.activeElement).toBe(box());
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it("Enter during IME composition sends nothing", async () => {
    type("mais curto");
    await act(async () => {
      fireEvent.keyDown(box(), { key: "Enter", isComposing: true });
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("zero-width-only and over-300 instructions cannot be sent", async () => {
    type("​ ​");
    expect(btn("Enviar alterações").disabled).toBe(true);
    await act(async () => {
      fireEvent.keyDown(box(), { key: "Enter" });
    });
    type("a".repeat(301)); // IME can exceed maxLength
    expect(btn("Enviar alterações").disabled).toBe(true);
    await act(async () => {
      fireEvent.keyDown(box(), { key: "Enter" });
    });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("ai_output_invalid on a revision shows the refinement hint and keeps the instruction", async () => {
    fetchMock.mockResolvedValue({
      ok: false,
      status: 502,
      json: async () => ({ ok: false, error: "ai_output_invalid" }),
    });
    type("acrescenta o link");
    await click("Enviar alterações");
    expect(screen.getByRole("alert").textContent).toMatch(/formulário de resposta/);
    expect(screen.queryByText(/Tente novamente/)).toBeNull();
    expect(box().value).toBe("acrescenta o link");
  });

  it("real ItemChanged during an in-flight refinement aborts it and the old response never shows", async () => {
    let resolve!: (v: unknown) => void;
    fetchMock.mockImplementation(
      (_u: string, init: { signal: AbortSignal }) =>
        new Promise((r, rej) => {
          resolve = r;
          init.signal.addEventListener("abort", () => rej(new Error("abort")));
        })
    );
    type("mais curto");
    await click("Enviar alterações");
    const signal = fetchMock.mock.calls[1][1].signal as AbortSignal;
    await act(async () => handler?.());
    expect(signal.aborted).toBe(true);
    await act(async () => resolve(res("VELHA")));
    expect(screen.queryByText("VELHA")).toBeNull();
    expect(screen.queryByText("Resposta atualizada")).toBeNull();
  });

  it("a suggestion over 2000 chars disables the box instead of truncating", async () => {
    fetchMock.mockResolvedValue(res("x".repeat(2001)));
    await click("Sugerir resposta");
    expect(screen.queryByPlaceholderText("Ex.: mais curto, pede o prazo")).toBeNull();
    expect(screen.getByText(/demasiado longa/)).toBeTruthy();
  });
});
