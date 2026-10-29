/* global Office */
import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

const T = "11111111-1111-1111-1111-111111111111";
const C = "22222222-2222-2222-2222-222222222222";
const URL_ = "https://abc.supabase.co/functions/v1/whoami";

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

const setNaa = (supported: boolean) => {
  (globalThis as any).Office = {
    context: { requirements: { isSetSupported: vi.fn(() => supported) } },
  };
};
const fetchMock = vi.fn();
const reply = (status: number, body: unknown = {}) =>
  fetchMock.mockResolvedValue({ status, json: async () => body });

beforeEach(() => {
  vi.clearAllMocks();
  h.configured.v = true;
  setNaa(true);
  h.create.mockResolvedValue({ acquireTokenSilent: h.silent, acquireTokenPopup: h.popup });
  h.silent.mockResolvedValue({ accessToken: "tok" });
  vi.stubGlobal("fetch", fetchMock);
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});

describe("BackendStatus", () => {
  // Must run first: auth.ts caches the MSAL app after the first creation.
  it("AC-2/3/6: creates NAA app with single-tenant authority, exact scopes, bearer-only GET", async () => {
    reply(200, { ok: true, name: "Ada" });
    render(<BackendStatus />);
    expect(await screen.findByText("Ligado como Ada")).toBeTruthy();
    expect(Office.context.requirements.isSetSupported).toHaveBeenCalledWith("NestedAppAuth", "1.1");
    expect(h.create).toHaveBeenCalledWith({
      auth: { clientId: C, authority: `https://login.microsoftonline.com/${T}` },
    });
    expect(h.silent).toHaveBeenCalledWith({ scopes: [`api://${C}/access_as_user`, "profile"] });
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe(URL_);
    expect(init.method).toBe("GET");
    expect(init.headers).toEqual({ Authorization: "Bearer tok" });
    expect(init.body).toBeUndefined();
    expect(init.signal).toBeInstanceOf(AbortSignal);
  });

  it("AC-1: unsupported wins over placeholder config", async () => {
    setNaa(false);
    h.configured.v = false;
    render(<BackendStatus />);
    expect(await screen.findByText(/não suporta o início de sessão/)).toBeTruthy();
  });

  it("AC-1: unsupported client shows the line, no MSAL, no fetch", async () => {
    setNaa(false);
    render(<BackendStatus />);
    expect(await screen.findByText(/não suporta o início de sessão/)).toBeTruthy();
    expect(h.create).not.toHaveBeenCalled();
    expect(h.silent).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("placeholder config: not-configured line, no MSAL, no fetch", async () => {
    h.configured.v = false;
    render(<BackendStatus />);
    expect(await screen.findByText(/ainda não está configurado/)).toBeTruthy();
    expect(h.silent).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("AC-4: interaction required shows Sign in; popup only on click", async () => {
    h.silent.mockRejectedValue(new h.InteractionRequiredAuthError());
    h.popup.mockResolvedValue({ accessToken: "tok2" });
    reply(200, { ok: true, name: null });
    render(<BackendStatus />);
    const btn = await screen.findByText("Iniciar sessão");
    expect(h.popup).not.toHaveBeenCalled();
    fireEvent.click(btn);
    expect(await screen.findByText("Ligado")).toBeTruthy();
    expect(h.popup).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0][1].headers).toEqual({ Authorization: "Bearer tok2" });
  });

  it("AC-4: other silent error shows the MSAL errorCode and Retry", async () => {
    h.silent.mockRejectedValue(Object.assign(new Error("x"), { errorCode: "network_error" }));
    render(<BackendStatus />);
    expect(await screen.findByText(/network_error/)).toBeTruthy();
    expect(screen.getByText("Tentar novamente")).toBeTruthy();
    expect(h.popup).not.toHaveBeenCalled();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("AC-5: popup failure shows Falha ao iniciar sessão with code, Retry, no backend call", async () => {
    h.silent.mockRejectedValue(new h.InteractionRequiredAuthError());
    h.popup.mockRejectedValue(Object.assign(new Error("x"), { errorCode: "popup_window_error" }));
    render(<BackendStatus />);
    fireEvent.click(await screen.findByText("Iniciar sessão"));
    expect(await screen.findByText("Falha ao iniciar sessão (popup_window_error)")).toBeTruthy();
    expect(screen.getByText("Tentar novamente")).toBeTruthy();
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("AC-7: 401 maps to rejected", async () => {
    reply(401, { ok: false });
    render(<BackendStatus />);
    expect(await screen.findByText("O backend rejeitou o início de sessão (401)")).toBeTruthy();
    expect(screen.getByText("Tentar novamente")).toBeTruthy();
  });

  it("AC-7: 500 and network error map to unreachable", async () => {
    reply(500);
    const { unmount } = render(<BackendStatus />);
    expect(await screen.findByText("Backend inacessível (500)")).toBeTruthy();
    unmount();
    fetchMock.mockRejectedValue(new TypeError("Failed to fetch"));
    render(<BackendStatus />);
    expect(await screen.findByText("Backend inacessível")).toBeTruthy();
  });

  it("AC-7: name is rendered as text, not HTML", async () => {
    reply(200, { ok: true, name: "<img src=x onerror=alert(1)>" });
    const { container } = render(<BackendStatus />);
    expect(await screen.findByText(/Ligado como <img/)).toBeTruthy();
    expect(container.querySelector("img")).toBeNull();
  });

  it("AC-6: request is aborted after 10 s", async () => {
    vi.useFakeTimers();
    try {
      fetchMock.mockImplementation(
        (_u: string, init: RequestInit) =>
          new Promise((_r, rej) =>
            init.signal!.addEventListener("abort", () => rej(new Error("aborted")))
          )
      );
      render(<BackendStatus />);
      await vi.advanceTimersByTimeAsync(9000);
      expect(fetchMock.mock.calls[0][1].signal.aborted).toBe(false);
      await vi.advanceTimersByTimeAsync(1001);
      expect(fetchMock.mock.calls[0][1].signal.aborted).toBe(true);
    } finally {
      vi.useRealTimers();
    }
    expect(await screen.findByText("Backend inacessível")).toBeTruthy();
  });
});
