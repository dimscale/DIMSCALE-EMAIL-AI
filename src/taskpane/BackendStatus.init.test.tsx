import * as React from "react";
import { afterEach, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";

const create = vi.hoisted(() => vi.fn());
vi.mock("@azure/msal-browser", () => ({
  InteractionRequiredAuthError: class extends Error {},
  createNestablePublicClientApplication: create,
}));
vi.mock("./config", () => ({
  CLIENT_ID: "c",
  AUTHORITY: "https://login.microsoftonline.com/t",
  SCOPES: ["s"],
  FUNCTION_URL: "https://abc.supabase.co/functions/v1/whoami",
  isConfigured: () => true,
}));

import BackendStatus from "./BackendStatus";

afterEach(cleanup);

// Own file: auth.ts caches the MSAL app per module instance.
it("MSAL creation failure shows an error with Retry, and Retry recreates the app", async () => {
  (globalThis as any).Office = { context: { requirements: { isSetSupported: () => true } } };
  create.mockRejectedValueOnce(new Error("boom")).mockResolvedValue({
    acquireTokenSilent: vi.fn().mockRejectedValue(new Error("x")),
  });
  render(<BackendStatus />);
  expect(await screen.findByRole("alert")).toBeTruthy();
  expect(screen.queryByText(/A ligar/)).toBeNull();
  fireEvent.click(screen.getByText("Tentar novamente"));
  await vi.waitFor(() => expect(create).toHaveBeenCalledTimes(2));
});
