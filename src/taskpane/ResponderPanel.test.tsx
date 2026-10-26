import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";

vi.mock("./auth", () => ({
  getTokenSilent: vi.fn(),
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

import { getTokenSilent } from "./auth";
import { getGraphTokenSilent, requestGraphConsent, fetchRecentSenders } from "./graph";
import { fetchSettings, saveSettings, deleteSettings } from "./settingsClient";
import ResponderPanel from "./ResponderPanel";

afterEach(cleanup);

describe("ResponderPanel (F-013/AC-32 to AC-35)", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    (globalThis as any).Office = {
      context: {
        mailbox: {
          userProfile: {
            emailAddress: "thiago.santos@dimscale.com",
          },
        },
      },
    };

    (getTokenSilent as any).mockResolvedValue({
      kind: "token",
      token: "mock_backend_token",
    });

    (fetchSettings as any).mockResolvedValue({
      kind: "ok",
      data: {
        enabled: true,
        status: "active",
        contacts: ["ana@cliente.pt"],
      },
    });

    (getGraphTokenSilent as any).mockResolvedValue({
      kind: "token",
      token: "mock_graph_token",
    });

    (fetchRecentSenders as any).mockResolvedValue([
      { address: "ana@cliente.pt", name: "Ana Silva" },
      { address: "rui@empresa.com", name: "Rui Costa" },
    ]);
  });

  it("loads initial settings and displays switch, status and contacts", async () => {
    render(<ResponderPanel />);
    expect(await screen.findByText("Contactos a responder (1)")).toBeTruthy();
    expect(screen.getByRole("switch", { name: "Respostas automáticas" })).toBeTruthy();
    expect(screen.getByText("Estado: Ativo")).toBeTruthy();
    expect(screen.getByLabelText("Ana Silva ana@cliente.pt")).toBeTruthy();
    expect(screen.getByLabelText("Rui Costa rui@empresa.com")).toBeTruthy();
    expect(screen.getByText("Só é criado um rascunho. Nada é enviado. O texto destes e-mails é processado por um serviço de IA externo.")).toBeTruthy();
  });

  it("prevents turning on switch with 0 contacts", async () => {
    (fetchSettings as any).mockResolvedValueOnce({
      kind: "ok",
      data: { enabled: false, status: "off", contacts: [] },
    });
    (fetchRecentSenders as any).mockResolvedValueOnce([]);

    render(<ResponderPanel />);
    const sw = await screen.findByRole("switch", { name: "Respostas automáticas" });
    expect(screen.getByText("Contactos a responder (0)")).toBeTruthy();

    fireEvent.click(sw);
    expect(screen.getByText("Escolha pelo menos um contacto.")).toBeTruthy();
  });

  it("adds and removes contacts via checkbox selection", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    const ruiCheckbox = screen.getByLabelText("Rui Costa rui@empresa.com");
    fireEvent.click(ruiCheckbox);

    expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();

    fireEvent.click(ruiCheckbox);
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
  });

  it("adds manual contact with validation", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    const input = screen.getByRole("textbox", { name: "Adicionar endereço" });
    const addButton = screen.getByRole("button", { name: "Adicionar" });

    // Invalid format
    fireEvent.change(input, { target: { value: "invalid-email" } });
    fireEvent.click(addButton);
    expect(screen.getByText("Endereço de email inválido.")).toBeTruthy();

    // Own address
    fireEvent.change(input, { target: { value: "thiago.santos@dimscale.com" } });
    fireEvent.click(addButton);
    expect(screen.getByText("Não pode adicionar o seu próprio endereço.")).toBeTruthy();

    // Already exists
    fireEvent.change(input, { target: { value: "ana@cliente.pt" } });
    fireEvent.click(addButton);
    expect(screen.getByText("Este endereço já está na lista.")).toBeTruthy();

    // Valid new address
    fireEvent.change(input, { target: { value: "novo@contacto.pt" } });
    fireEvent.click(addButton);
    expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
    expect(screen.getByLabelText("novo@contacto.pt")).toBeTruthy();
  });

  it("disables unticked checkboxes when cap of 50 is reached", async () => {
    const fiftyContacts = Array.from({ length: 50 }, (_, i) => `user${i}@exemplo.pt`);
    (fetchSettings as any).mockResolvedValueOnce({
      kind: "ok",
      data: { enabled: true, status: "active", contacts: fiftyContacts },
    });
    (fetchRecentSenders as any).mockResolvedValueOnce([
      ...fiftyContacts.map((c) => ({ address: c, name: "" })),
      { address: "extra@exemplo.pt", name: "Extra" },
    ]);

    render(<ResponderPanel />);
    expect(await screen.findByText("Contactos a responder (50)")).toBeTruthy();
    expect(screen.getByText("Máximo de 50 contactos.")).toBeTruthy();

    const extraBox = screen.getByLabelText("Extra extra@exemplo.pt") as HTMLInputElement;
    expect(extraBox.disabled).toBe(true);
  });

  it("shows Permitir acesso button when consent is needed", async () => {
    (getGraphTokenSilent as any).mockResolvedValueOnce({ kind: "needConsent" });
    (requestGraphConsent as any).mockResolvedValueOnce({ kind: "token", token: "new_token" });

    render(<ResponderPanel />);
    const allowBtn = await screen.findByRole("button", { name: "Permitir acesso" });
    expect(screen.getByText("Permitir acesso aos remetentes recentes")).toBeTruthy();

    fireEvent.click(allowBtn);
    await waitFor(() => {
      expect(requestGraphConsent).toHaveBeenCalled();
    });
  });

  it("saves settings on Guardar click", async () => {
    (saveSettings as any).mockResolvedValue({
      kind: "ok",
      data: { enabled: true, status: "pending", contacts: ["ana@cliente.pt"] },
    });

    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    const saveBtn = screen.getByRole("button", { name: "Guardar" });
    fireEvent.click(saveBtn);

    await waitFor(() => {
      expect(saveSettings).toHaveBeenCalledWith(
        "mock_backend_token",
        expect.objectContaining({
          enabled: true,
          ownAddress: "thiago.santos@dimscale.com",
          contacts: ["ana@cliente.pt"],
        })
      );
      expect(screen.getByText("Definições guardadas.")).toBeTruthy();
      expect(screen.getByText("Estado: A ativar…")).toBeTruthy();
    });
  });

  it("handles Apagar tudo with confirmation", async () => {
    (deleteSettings as any).mockResolvedValue({ ok: true });

    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    const deleteBtn = screen.getByRole("button", { name: "Apagar tudo" });
    fireEvent.click(deleteBtn);

    expect(screen.getByText("Tem a certeza que deseja apagar todos os contactos?")).toBeTruthy();

    const confirmYesBtn = screen.getByRole("button", { name: "Sim, apagar" });
    fireEvent.click(confirmYesBtn);

    await waitFor(() => {
      expect(deleteSettings).toHaveBeenCalledWith("mock_backend_token");
      expect(screen.getByText("Estado: Desligado")).toBeTruthy();
      expect(screen.getByText("Contactos a responder (0)")).toBeTruthy();
    });
  });

  it("renders display names and addresses literally (no HTML injection)", async () => {
    (fetchRecentSenders as any).mockResolvedValueOnce([
      { address: "injection@test.pt", name: "<script>alert('xss')</script>" },
    ]);

    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    expect(screen.getByText("<script>alert('xss')</script> <injection@test.pt>")).toBeTruthy();
  });
});
