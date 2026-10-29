// QA (independent, F-013 step 10b): the Responder tab makes no Graph or settings call until it is opened.
import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

vi.mock("./auth", () => ({
  getTokenSilent: vi.fn(),
  getApp: vi.fn(),
  signIn: vi.fn(),
  isNaaSupported: vi.fn(() => true),
}));

vi.mock("./graph", () => ({
  getGraphTokenSilent: vi.fn(),
  requestGraphConsent: vi.fn(),
  fetchRecentSenders: vi.fn(),
  GRAPH_SCOPES: ["Mail.ReadBasic"],
}));

vi.mock("./settingsClient", () => ({
  fetchSettings: vi.fn(),
  saveSettings: vi.fn(),
  deleteSettings: vi.fn(),
}));

// Stub the Analisar-side and backend pieces so only the Responder tab's calls are observed.
vi.mock("./BackendStatus", () => ({ default: () => null }));
vi.mock("./MessagePanel", () => ({ default: () => null }));

import { getTokenSilent } from "./auth";
import { fetchRecentSenders, getGraphTokenSilent } from "./graph";
import { fetchSettings } from "./settingsClient";
import App from "./App";

const flush = () => new Promise((r) => setTimeout(r, 30));

beforeEach(() => {
  vi.clearAllMocks();
  (globalThis as any).Office = {
    CoercionType: { Text: "text" },
    context: {
      mailbox: { item: undefined, userProfile: { emailAddress: "thiago.santos@dimscale.com" } },
    },
  };
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_backend_token" });
  (fetchSettings as any).mockResolvedValue({
    kind: "ok",
    data: { enabled: true, status: "active", contacts: ["ana@cliente.pt"] },
  });
  (getGraphTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_graph_token" });
  (fetchRecentSenders as any).mockResolvedValue([{ address: "ana@cliente.pt", name: "Ana Silva" }]);
});

afterEach(cleanup);

describe("F-013 Responder tab is lazy (QA, mocked)", () => {
  it("no Graph or settings call happens while the Analisar tab is shown", async () => {
    render(<App />);
    await flush();
    expect(screen.getByRole("tab", { name: "Analisar" }).getAttribute("aria-selected")).toBe(
      "true"
    );
    expect(getTokenSilent).not.toHaveBeenCalled();
    expect(getGraphTokenSilent).not.toHaveBeenCalled();
    expect(fetchSettings).not.toHaveBeenCalled();
    expect(fetchRecentSenders).not.toHaveBeenCalled();
  });

  it("opening the Responder tab triggers the load once, and switching tabs back and forth does not reload", async () => {
    render(<App />);
    await flush();
    fireEvent.click(screen.getByRole("tab", { name: "Responder" }));
    expect(await screen.findByText("Contactos a responder (1)")).toBeTruthy();
    expect(getGraphTokenSilent).toHaveBeenCalledTimes(1);
    expect(fetchSettings).toHaveBeenCalledTimes(1);
    expect(fetchRecentSenders).toHaveBeenCalledTimes(1);

    fireEvent.click(screen.getByRole("tab", { name: "Analisar" }));
    fireEvent.click(screen.getByRole("tab", { name: "Responder" }));
    await flush();
    expect(fetchSettings).toHaveBeenCalledTimes(1);
    expect(getGraphTokenSilent).toHaveBeenCalledTimes(1);
    expect(fetchRecentSenders).toHaveBeenCalledTimes(1);
  });
});
