// QA (independent, F-013 step 10b): Responder panel states with mocked auth, Graph and settings client.
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
import { fetchRecentSenders, getGraphTokenSilent, requestGraphConsent } from "./graph";
import { deleteSettings, fetchSettings, saveSettings } from "./settingsClient";
import ResponderPanel from "./ResponderPanel";

const OWN = "thiago.santos@dimscale.com";
const CONSENT_BUTTON = "Autorizar acesso";

const ok = (enabled: boolean, status: string, contacts: string[]) => ({
  kind: "ok",
  data: { enabled, status, contacts },
});

const flush = () => new Promise((r) => setTimeout(r, 30));

beforeEach(() => {
  vi.clearAllMocks();
  (globalThis as any).Office = { context: { mailbox: { userProfile: { emailAddress: OWN } } } };
  (getTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_backend_token" });
  (fetchSettings as any).mockResolvedValue(ok(true, "active", ["ana@cliente.pt"]));
  (getGraphTokenSilent as any).mockResolvedValue({ kind: "token", token: "mock_graph_token" });
  (fetchRecentSenders as any).mockResolvedValue([
    { address: "ana@cliente.pt", name: "Ana Silva" },
    { address: "rui@empresa.com", name: "Rui Costa" },
  ]);
});

afterEach(cleanup);

describe("F-013/AC-34, AC-35 Responder panel (QA, mocked)", () => {
  it("shows 'Erro, a tentar novamente' when the settings function reports status error", async () => {
    (fetchSettings as any).mockResolvedValueOnce(ok(true, "error", ["ana@cliente.pt"]));
    render(<ResponderPanel />);
    expect(await screen.findByText("Estado: Erro, a tentar novamente")).toBeTruthy();
  });

  it("shows 'Conta fora do grupo piloto' for not_in_scope", async () => {
    (fetchSettings as any).mockResolvedValueOnce(ok(true, "not_in_scope", ["ana@cliente.pt"]));
    render(<ResponderPanel />);
    expect(await screen.findByText("Estado: Conta fora do grupo piloto")).toBeTruthy();
  });

  it("settings load failure shows an error with Tentar novamente, and every write control is disabled", async () => {
    (fetchSettings as any).mockResolvedValueOnce({ kind: "error" });
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (0)");
    expect(screen.getByText("Não foi possível carregar as definições.")).toBeTruthy();
    expect(
      (screen.getByRole("switch", { name: "Respostas automáticas" }) as HTMLInputElement).disabled
    ).toBe(true);
    expect((screen.getByRole("button", { name: "Guardar" }) as HTMLButtonElement).disabled).toBe(
      true
    );
    expect(
      (screen.getByRole("button", { name: "Apagar tudo" }) as HTMLButtonElement).disabled
    ).toBe(true);
    expect((screen.getByRole("button", { name: "Adicionar" }) as HTMLButtonElement).disabled).toBe(
      true
    );
    expect(
      (screen.getByRole("textbox", { name: "Adicionar endereço" }) as HTMLInputElement).disabled
    ).toBe(true);
  });

  it("after a failed load, clicking Guardar or the disabled controls sends no PUT and no DELETE", async () => {
    (fetchSettings as any).mockResolvedValueOnce({ kind: "error" });
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (0)");
    fireEvent.click(screen.getByRole("button", { name: "Guardar" }));
    fireEvent.click(screen.getByRole("button", { name: "Apagar tudo" }));
    fireEvent.click(screen.getByRole("switch", { name: "Respostas automáticas" }));
    await flush();
    expect(saveSettings).not.toHaveBeenCalled();
    expect(deleteSettings).not.toHaveBeenCalled();
  });

  it("Tentar novamente after a failed load reloads settings and then enables saving", async () => {
    (fetchSettings as any)
      .mockResolvedValueOnce({ kind: "error" })
      .mockResolvedValueOnce(ok(true, "active", ["ana@cliente.pt"]));
    render(<ResponderPanel />);
    fireEvent.click(await screen.findByRole("button", { name: "Tentar novamente" }));
    expect(await screen.findByText("Contactos a responder (1)")).toBeTruthy();
    expect(screen.queryByText("Não foi possível carregar as definições.")).toBeNull();
    expect((screen.getByRole("button", { name: "Guardar" }) as HTMLButtonElement).disabled).toBe(
      false
    );
    expect(fetchSettings).toHaveBeenCalledTimes(2);
  });

  it("settings 401 (rejected) is a load failure too", async () => {
    (fetchSettings as any).mockResolvedValueOnce({ kind: "rejected" });
    render(<ResponderPanel />);
    expect(await screen.findByText("Não foi possível carregar as definições.")).toBeTruthy();
    expect((screen.getByRole("button", { name: "Guardar" }) as HTMLButtonElement).disabled).toBe(
      true
    );
  });

  it("consent popup cancelled: shows the permission text, makes no listing call, manual entry still works", async () => {
    (getGraphTokenSilent as any).mockResolvedValueOnce({ kind: "needConsent" });
    (requestGraphConsent as any).mockResolvedValueOnce({ kind: "denied", code: "user_cancelled" });

    render(<ResponderPanel />);
    const allow = await screen.findByRole("button", { name: CONSENT_BUTTON });
    // silent first: no popup before the click
    expect(requestGraphConsent).not.toHaveBeenCalled();

    fireEvent.click(allow);
    expect(
      await screen.findByText(
        "Sem autorização para listar remetentes recentes. Pode adicionar endereços à mão."
      )
    ).toBeTruthy();
    expect(requestGraphConsent).toHaveBeenCalledTimes(1);
    expect(fetchRecentSenders).not.toHaveBeenCalled();
    expect(screen.queryByRole("button", { name: CONSENT_BUTTON })).toBeNull();

    fireEvent.change(screen.getByRole("textbox", { name: "Adicionar endereço" }), {
      target: { value: "novo@contacto.pt" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Adicionar" }));
    expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
  });

  it("denied silent Graph token (admin approval needed) shows the permission text and no candidates", async () => {
    (getGraphTokenSilent as any).mockResolvedValueOnce({
      kind: "denied",
      code: "consent_required",
    });
    render(<ResponderPanel />);
    expect(
      await screen.findByText(
        "Sem autorização para listar remetentes recentes. Pode adicionar endereços à mão."
      )
    ).toBeTruthy();
    expect(fetchRecentSenders).not.toHaveBeenCalled();
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
  });

  it("the active prop gates loading: nothing is requested while inactive, and loading starts on first activation only", async () => {
    const { rerender } = render(<ResponderPanel active={false} />);
    await flush();
    expect(getTokenSilent).not.toHaveBeenCalled();
    expect(fetchSettings).not.toHaveBeenCalled();
    expect(getGraphTokenSilent).not.toHaveBeenCalled();
    expect(fetchRecentSenders).not.toHaveBeenCalled();

    rerender(<ResponderPanel active={true} />);
    expect(await screen.findByText("Contactos a responder (1)")).toBeTruthy();
    expect(fetchSettings).toHaveBeenCalledTimes(1);
    expect(fetchRecentSenders).toHaveBeenCalledTimes(1);

    rerender(<ResponderPanel active={false} />);
    rerender(<ResponderPanel active={true} />);
    await flush();
    expect(fetchSettings).toHaveBeenCalledTimes(1);
    expect(getGraphTokenSilent).toHaveBeenCalledTimes(1);
  });

  it.each([
    ["429", "Graph request failed with status 429"],
    ["500", "Graph request failed with status 500"],
    ["network", "Failed to fetch"],
  ])(
    "Graph listing failure (%s) keeps the pane usable with saved contacts and manual entry",
    async (_label, msg) => {
      (fetchRecentSenders as any).mockRejectedValueOnce(new Error(msg));
      render(<ResponderPanel />);
      expect(await screen.findByText("Contactos a responder (1)")).toBeTruthy();
      await waitFor(() => expect(screen.queryByText("A obter remetentes recentes…")).toBeNull());
      expect(screen.getByLabelText("ana@cliente.pt")).toBeTruthy();
      expect(
        screen.getByText("Não foi possível obter os remetentes recentes. Pode adicionar endereços manualmente.")
      ).toBeTruthy();
      expect(screen.queryByText("Nenhum remetente recente encontrado nos últimos 30 dias.")).toBeNull();

      fireEvent.change(screen.getByRole("textbox", { name: "Adicionar endereço" }), {
        target: { value: "novo@contacto.pt" },
      });
      fireEvent.click(screen.getByRole("button", { name: "Adicionar" }));
      expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
    }
  );

  it.each([
    ["422 from the server", { kind: "error", status: 422 }, "Os dados não foram aceites. Verifique os endereços."],
    ["500 from the server", { kind: "error", status: 500 }, "Erro de ligação"],
    ["network failure", { kind: "error" }, "Erro de ligação"],
    ["401 (rejected)", { kind: "rejected" }, "Erro de ligação"],
  ])(
    "save failure (%s) shows the matching error, keeps the local selection and never shows success",
    async (_label, result, expectedText) => {
      (saveSettings as any).mockResolvedValueOnce(result);
      render(<ResponderPanel />);
      await screen.findByText("Contactos a responder (1)");

      fireEvent.click(screen.getByRole("checkbox", { name: "Rui Costa rui@empresa.com" }));
      fireEvent.click(screen.getByRole("button", { name: "Guardar" }));

      expect(await screen.findByText(expectedText)).toBeTruthy();
      expect(screen.queryByText("Definições guardadas.")).toBeNull();
      expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
      expect(screen.getByText("Estado: Ativo")).toBeTruthy();
    }
  );

  it("Guardar sends exactly {enabled, ownAddress, contacts}: no display names, selected addresses lower-cased", async () => {
    (saveSettings as any).mockResolvedValueOnce(
      ok(true, "pending", ["ana@cliente.pt", "rui@empresa.com"])
    );
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    fireEvent.click(screen.getByRole("checkbox", { name: "Rui Costa rui@empresa.com" }));
    fireEvent.click(screen.getByRole("button", { name: "Guardar" }));

    await waitFor(() => expect(saveSettings).toHaveBeenCalledTimes(1));
    expect(saveSettings).toHaveBeenCalledWith("mock_backend_token", {
      enabled: true,
      ownAddress: OWN,
      contacts: ["ana@cliente.pt", "rui@empresa.com"],
    });
  });

  it("switch off with 0 contacts: Guardar saves enabled false and an empty list, and Apagar tudo is disabled", async () => {
    (fetchSettings as any).mockResolvedValueOnce(ok(false, "off", []));
    (fetchRecentSenders as any).mockResolvedValueOnce([]);
    (saveSettings as any).mockResolvedValueOnce(ok(false, "off", []));

    render(<ResponderPanel />);
    const sw = (await screen.findByRole("switch", {
      name: "Respostas automáticas",
    })) as HTMLInputElement;
    expect(sw.checked).toBe(false);
    expect(sw.disabled).toBe(false);
    expect(
      (screen.getByRole("button", { name: "Apagar tudo" }) as HTMLButtonElement).disabled
    ).toBe(true);

    fireEvent.click(screen.getByRole("button", { name: "Guardar" }));
    await waitFor(() =>
      expect(saveSettings).toHaveBeenCalledWith("mock_backend_token", {
        enabled: false,
        ownAddress: OWN,
        contacts: [],
      })
    );
    expect(await screen.findByText("Definições guardadas.")).toBeTruthy();
  });

  it("the 50 cap at the boundary: 49 allows the 50th tick, then the 51st candidate is disabled and manual add is blocked", async () => {
    const forty9 = Array.from({ length: 49 }, (_, i) => `user${i}@exemplo.pt`);
    (fetchSettings as any).mockResolvedValueOnce(ok(true, "active", forty9));
    (fetchRecentSenders as any).mockResolvedValueOnce([
      ...forty9.map((a) => ({ address: a, name: "" })),
      { address: "c50@exemplo.pt", name: "Cinquenta" },
      { address: "c51@exemplo.pt", name: "Cinquenta e um" },
    ]);

    render(<ResponderPanel />);
    expect(await screen.findByText("Contactos a responder (49)")).toBeTruthy();
    expect(screen.queryByText("Máximo de 50 contactos.")).toBeNull();

    const c50 = screen.getByLabelText("Cinquenta c50@exemplo.pt") as HTMLInputElement;
    const c51 = screen.getByLabelText("Cinquenta e um c51@exemplo.pt") as HTMLInputElement;
    expect(c50.disabled).toBe(false);
    fireEvent.click(c50);
    expect(screen.getByText("Contactos a responder (50)")).toBeTruthy();
    expect(screen.getByText("Máximo de 50 contactos.")).toBeTruthy();
    expect(c51.disabled).toBe(true);

    fireEvent.click(c51);
    expect(screen.getByText("Contactos a responder (50)")).toBeTruthy();

    const addBtn = screen.getByRole("button", { name: "Adicionar" }) as HTMLButtonElement;
    fireEvent.change(screen.getByRole("textbox", { name: "Adicionar endereço" }), {
      target: { value: "extra@x.pt" },
    });
    expect(addBtn.disabled).toBe(true);
  });

  it("manual add normalises case and whitespace, rejects duplicates and adds once per click", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    const input = screen.getByRole("textbox", { name: "Adicionar endereço" });
    const addBtn = screen.getByRole("button", { name: "Adicionar" });

    fireEvent.change(input, { target: { value: "   ANA@Cliente.PT   " } });
    fireEvent.click(addBtn);
    expect(screen.getByText("Este endereço já está na lista.")).toBeTruthy();
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();

    fireEvent.change(input, { target: { value: "  Novo@Contacto.PT  " } });
    fireEvent.click(addBtn);
    expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
    expect(screen.getByLabelText("novo@contacto.pt")).toBeTruthy();

    fireEvent.change(input, { target: { value: "NOVO@contacto.pt" } });
    fireEvent.click(addBtn);
    expect(screen.getByText("Este endereço já está na lista.")).toBeTruthy();

    fireEvent.change(input, { target: { value: "  THIAGO.SANTOS@DIMSCALE.COM " } });
    fireEvent.click(addBtn);
    expect(screen.getByText("Não pode adicionar o seu próprio endereço.")).toBeTruthy();

    for (const bad of ["a b@x.pt", "sem-arroba", "a@b", "@x.pt"]) {
      fireEvent.change(input, { target: { value: bad } });
      fireEvent.click(addBtn);
      expect(screen.getByText("Endereço de email inválido.")).toBeTruthy();
    }

    fireEvent.change(input, { target: { value: "    " } });
    expect((addBtn as HTMLButtonElement).disabled).toBe(true);
  });

  it("manual add of a candidate that is not selected selects it (no duplicate row)", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    const input = screen.getByRole("textbox", { name: "Adicionar endereço" });
    fireEvent.change(input, { target: { value: " RUI@empresa.com " } });
    fireEvent.click(screen.getByRole("button", { name: "Adicionar" }));

    expect(screen.getByText("Contactos a responder (2)")).toBeTruthy();
    const box = screen.getByRole("checkbox", {
      name: "Rui Costa rui@empresa.com",
    }) as HTMLInputElement;
    expect(box.checked).toBe(true);
    expect(screen.getAllByRole("checkbox")).toHaveLength(2);
  });

  it("Apagar tudo cancel path makes no DELETE and no PUT, and keeps the list", async () => {
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    fireEvent.click(screen.getByRole("button", { name: "Apagar tudo" }));
    expect(screen.getByText("Tem a certeza que deseja apagar todos os contactos?")).toBeTruthy();
    fireEvent.click(screen.getByRole("button", { name: "Cancelar" }));

    expect(screen.queryByText("Tem a certeza que deseja apagar todos os contactos?")).toBeNull();
    expect(screen.getByRole("button", { name: "Guardar" })).toBeTruthy();
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
    await flush();
    expect(deleteSettings).not.toHaveBeenCalled();
    expect(saveSettings).not.toHaveBeenCalled();
  });

  it("Apagar tudo confirmed clears the list and shows the deleted text", async () => {
    (deleteSettings as any).mockResolvedValueOnce({ ok: true });
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    fireEvent.click(screen.getByRole("button", { name: "Apagar tudo" }));
    fireEvent.click(screen.getByRole("button", { name: "Sim, apagar" }));
    expect(await screen.findByText("Contactos apagados.")).toBeTruthy();
    expect(screen.queryByText("Definições guardadas.")).toBeNull();
    expect(deleteSettings).toHaveBeenCalledTimes(1);
    expect(screen.getByText("Estado: Desligado")).toBeTruthy();
    expect(screen.getByText("Contactos a responder (0)")).toBeTruthy();
  });

  it("Apagar tudo confirmed with a failing DELETE shows an error and keeps the list", async () => {
    (deleteSettings as any).mockResolvedValueOnce({ ok: false });
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    fireEvent.click(screen.getByRole("button", { name: "Apagar tudo" }));
    fireEvent.click(screen.getByRole("button", { name: "Sim, apagar" }));
    expect(await screen.findByText("Erro de ligação")).toBeTruthy();
    expect(screen.getByText("Contactos a responder (1)")).toBeTruthy();
  });

  it("names with HTML and RTL controls render as text, never as markup", async () => {
    const rtl = "‮liame‬";
    (fetchRecentSenders as any).mockResolvedValueOnce([
      { address: "img@teste.pt", name: `<img src=x onerror="alert(1)"> ${rtl}` },
      { address: "heb@teste.pt", name: "שלום <b>mundo</b> <script>x()</script>" },
    ]);
    const { container } = render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");

    expect(container.querySelector("img")).toBeNull();
    expect(container.querySelector("b")).toBeNull();
    expect(container.querySelector("script")).toBeNull();
    expect(container.textContent).toContain('<img src=x onerror="alert(1)">');
    expect(container.textContent).toContain("<b>mundo</b>");
    expect(container.textContent).toContain("<script>x()</script>");
    expect(container.querySelectorAll('[dir="auto"]').length).toBeGreaterThanOrEqual(2);
  });

  it("no browser storage is written during the flow, and no address or token reaches the console", async () => {
    const setItem = vi.spyOn(Storage.prototype, "setItem");
    const logs: unknown[][] = [];
    const spies = (["log", "info", "warn", "error", "debug"] as const).map((m) =>
      vi.spyOn(console, m).mockImplementation((...a: unknown[]) => {
        logs.push(a);
      })
    );

    (saveSettings as any).mockResolvedValueOnce({ kind: "error", status: 422 });
    render(<ResponderPanel />);
    await screen.findByText("Contactos a responder (1)");
    fireEvent.click(screen.getByRole("checkbox", { name: "Rui Costa rui@empresa.com" }));
    fireEvent.click(screen.getByRole("button", { name: "Guardar" }));
    await screen.findByText("Os dados não foram aceites. Verifique os endereços.");

    expect(setItem).not.toHaveBeenCalled();
    const flat = JSON.stringify(logs);
    expect(flat).not.toContain("ana@cliente.pt");
    expect(flat).not.toContain("rui@empresa.com");
    expect(flat).not.toContain("mock_backend_token");
    spies.forEach((s) => s.mockRestore());
    setItem.mockRestore();
  });
});
