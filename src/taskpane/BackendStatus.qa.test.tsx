/* global Office */
// QA additions for F-002 (independent check). MSAL, fetch and Office are mocked; no live calls.
import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

const T = "11111111-1111-1111-1111-111111111111";
const C = "22222222-2222-2222-2222-222222222222";
const URL_ = "https://abc.supabase.co/functions/v1/whoami";
const TOKEN = "eyJhbGciOi.SECRET-TOKEN-VALUE.sig";
const NAME = "Ada Lovelace";

const h = vi.hoisted(() => {
  class InteractionRequiredAuthError extends Error {
    errorCode = "interaction_required";
  }
  return {
    InteractionRequiredAuthError,
    silent: vi.fn(),
    popup: vi.fn(),
    create: vi.fn(),
    configured: { v: true },
  };
});

vi.mock("@azure/msal-browser", () => ({
  InteractionRequiredAuthError: h.InteractionRequiredAuthError,
  createNestablePublicClientApplication: h.create,
}));
vi.mock("./config", () => ({
  TENANT_ID: "11111111-1111-1111-1111-111111111111",
  CLIENT_ID: "22222222-2222-2222-2222-222222222222",
  FUNCTION_URL: "https://abc.supabase.co/functions/v1/whoami",
  AUTHORITY: "https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111",
  SCOPES: ["api://22222222-2222-2222-2222-222222222222/access_as_user", "profile"],
  isConfigured: () => h.configured.v,
}));

import BackendStatus from "./BackendStatus";
import MessagePanel from "./MessagePanel";
import { strings } from "./strings";

const fetchMock = vi.fn();
const ok = (name: unknown) => ({ status: 200, json: async () => ({ ok: true, name }) });
const status = (code: number) => ({ status: code, json: async () => ({ ok: false }) });

const setOffice = (opts: { naa?: boolean; item?: unknown }) => {
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: {
      mailbox: { item: opts.item },
      requirements: { isSetSupported: vi.fn(() => opts.naa ?? true) },
    },
  };
};

const makeItem = () => ({
  itemId: "AAMk-ITEM-ID-123",
  subject: "Quarterly plan",
  from: { displayName: "Ann", emailAddress: "ann@example.com" },
  dateTimeCreated: new Date(2026, 0, 2, 3, 4),
  body: { getAsync: (_t: string, cb: (r: { status: string; value?: string }) => void) => cb({ status: "succeeded", value: "SECRET BODY TEXT" }) },
});

beforeEach(() => {
  vi.clearAllMocks();
  fetchMock.mockReset();
  h.silent.mockReset();
  h.popup.mockReset();
  h.configured.v = true;
  setOffice({ naa: true });
  h.create.mockResolvedValue({ acquireTokenSilent: h.silent, acquireTokenPopup: h.popup });
  h.silent.mockResolvedValue({ accessToken: "tok" });
  vi.stubGlobal("fetch", fetchMock);
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("BackendStatus QA (mocked MSAL, fetch and Office)", () => {
  it("Retry after an unreachable backend calls MSAL and the backend again, then connects", async () => {
    fetchMock.mockRejectedValueOnce(new TypeError("Failed to fetch")).mockResolvedValueOnce(ok("Ada"));
    render(<BackendStatus />);
    expect(await screen.findByText("Backend inacessível")).toBeTruthy();
    fireEvent.click(screen.getByText("Tentar novamente"));
    expect(await screen.findByText("Ligado como Ada")).toBeTruthy();
    expect(h.silent).toHaveBeenCalledTimes(2);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[1][1].headers).toEqual({ Authorization: "Bearer tok" });
  });

  it("Retry after a 401 asks MSAL for a token again and sends the fresh one", async () => {
    h.silent.mockResolvedValueOnce({ accessToken: "old" }).mockResolvedValueOnce({ accessToken: "new" });
    fetchMock.mockResolvedValueOnce(status(401)).mockResolvedValueOnce(ok(null));
    render(<BackendStatus />);
    expect(await screen.findByText("O backend rejeitou o início de sessão (401)")).toBeTruthy();
    fireEvent.click(screen.getByText("Tentar novamente"));
    expect(await screen.findByText("Ligado ao backend")).toBeTruthy();
    expect(fetchMock.mock.calls[0][1].headers).toEqual({ Authorization: "Bearer old" });
    expect(fetchMock.mock.calls[1][1].headers).toEqual({ Authorization: "Bearer new" });
  });

  it("Retry after a failed popup starts silently and never opens a second popup by itself", async () => {
    h.silent.mockRejectedValue(new h.InteractionRequiredAuthError());
    h.popup.mockRejectedValue(Object.assign(new Error("closed"), { errorCode: "user_cancelled" }));
    render(<BackendStatus />);
    fireEvent.click(await screen.findByText("Iniciar sessão"));
    expect(await screen.findByText("Falha ao iniciar sessão (user_cancelled)")).toBeTruthy();
    fireEvent.click(screen.getByText("Tentar novamente"));
    expect(await screen.findByText("Inicie sessão para ligar ao backend DIMSCALE")).toBeTruthy();
    expect(h.silent).toHaveBeenCalledTimes(2);
    expect(h.popup).toHaveBeenCalledTimes(1);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("sign-in popup is requested with exactly the spec scopes", async () => {
    h.silent.mockRejectedValue(new h.InteractionRequiredAuthError());
    h.popup.mockResolvedValue({ accessToken: "tok2" });
    fetchMock.mockResolvedValue(ok(null));
    render(<BackendStatus />);
    fireEvent.click(await screen.findByText("Iniciar sessão"));
    expect(await screen.findByText("Ligado ao backend")).toBeTruthy();
    expect(h.popup).toHaveBeenCalledWith({ scopes: [`api://${C}/access_as_user`, "profile"] });
  });

  it("nothing from the token, the name or the MSAL message reaches any console method", async () => {
    const spies = ["log", "info", "warn", "error", "debug"].map((m) => vi.spyOn(console, m as "log").mockImplementation(() => {}));
    try {
      // 1: success with a name and a token
      h.silent.mockResolvedValue({ accessToken: TOKEN });
      fetchMock.mockResolvedValueOnce(ok(NAME));
      const first = render(<BackendStatus />);
      expect(await screen.findByText(`Ligado como ${NAME}`)).toBeTruthy();
      first.unmount();
      // 2: 401 for the same token
      fetchMock.mockResolvedValueOnce(status(401));
      const second = render(<BackendStatus />);
      expect(await screen.findByText("O backend rejeitou o início de sessão (401)")).toBeTruthy();
      second.unmount();
      // 3: network failure whose message carries the token, and an MSAL error whose message carries it too
      fetchMock.mockRejectedValueOnce(new TypeError(`fetch ${TOKEN}`));
      const third = render(<BackendStatus />);
      expect(await screen.findByText("Backend inacessível")).toBeTruthy();
      third.unmount();
      h.silent.mockRejectedValueOnce(Object.assign(new Error(`leak ${TOKEN}`), { errorCode: "x_code" }));
      render(<BackendStatus />);
      expect(await screen.findByText(/x_code/)).toBeTruthy();
      const calls = spies.flatMap((s) => s.mock.calls);
      for (const args of calls) {
        const text = JSON.stringify(args);
        expect(text).not.toContain(TOKEN);
        expect(text).not.toContain("SECRET-TOKEN-VALUE");
        expect(text).not.toContain(NAME);
      }
    } finally {
      spies.forEach((s) => s.mockRestore());
    }
  });

  it("the request is a bare GET: only the Authorization header, no body, no query, no email fields", async () => {
    setOffice({ naa: true, item: makeItem() });
    fetchMock.mockResolvedValue(ok(null));
    render(
      <>
        <MessagePanel />
        <BackendStatus />
      </>
    );
    expect(await screen.findByText("Ligado ao backend")).toBeTruthy();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(URL_);
    expect(Object.keys(init).sort()).toEqual(["headers", "method", "signal"]);
    expect(init.method).toBe("GET");
    expect(Object.keys(init.headers)).toEqual(["Authorization"]);
    const sent = JSON.stringify([url, init.headers]);
    for (const leak of ["Quarterly plan", "ann@example.com", "Ann", "SECRET BODY TEXT", "AAMk-ITEM-ID-123", "subject", "itemId"]) {
      expect(sent).not.toContain(leak);
    }
  });

  it("F-001 message context still renders when the backend is unreachable", async () => {
    setOffice({ naa: true, item: makeItem() });
    fetchMock.mockRejectedValue(new TypeError("Failed to fetch"));
    render(
      <>
        <MessagePanel />
        <BackendStatus />
      </>
    );
    expect(await screen.findByText("Backend inacessível")).toBeTruthy();
    expect(screen.getByText("Quarterly plan")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "Mensagem" }));
    expect(screen.getByText("SECRET BODY TEXT")).toBeTruthy();
  });

  it("F-001 message context still renders on an unsupported client, with no MSAL or fetch", async () => {
    setOffice({ naa: false, item: makeItem() });
    render(
      <>
        <MessagePanel />
        <BackendStatus />
      </>
    );
    expect(await screen.findByText(/não suporta o início de sessão/)).toBeTruthy();
    expect(screen.getByText("Quarterly plan")).toBeTruthy();
    expect(h.create).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("a 200 with a non-JSON body is unreachable, never connected", async () => {
    fetchMock.mockResolvedValue({ status: 200, json: async () => { throw new SyntaxError("Unexpected token <"); } });
    render(<BackendStatus />);
    expect(await screen.findByText("Backend inacessível")).toBeTruthy();
    expect(screen.queryByText(/Ligado ao backend/)).toBeNull();
  });

  it("a 200 with ok:false is unreachable, never connected", async () => {
    fetchMock.mockResolvedValue({ status: 200, json: async () => ({ ok: false, name: "Ada" }) });
    render(<BackendStatus />);
    expect(await screen.findByText("Backend inacessível (200)")).toBeTruthy();
    expect(screen.queryByText(/Ligado ao backend/)).toBeNull();
  });

  it.each([["" as unknown], [null], [42], [{ first: "Ada" }]])("name %j is not rendered: shows plain 'Ligado ao backend'", async (name) => {
    fetchMock.mockResolvedValue(ok(name));
    render(<BackendStatus />);
    expect(await screen.findByText("Ligado ao backend")).toBeTruthy();
  });

  it("an MSAL error without errorCode shows unknown_error, not a blank line", async () => {
    h.silent.mockRejectedValue(new Error("x"));
    render(<BackendStatus />);
    expect(await screen.findByText("Não foi possível obter o token de início de sessão (unknown_error)")).toBeTruthy();
    expect(screen.getByText("Tentar novamente")).toBeTruthy();
  });

  it("Behavior: the working line uses the spec's ellipsis (U+2026)", () => {
    expect(strings.backendWorking).toBe("A ligar ao backend DIMSCALE…");
  });
});
