/* global Office */

import { describe, expect, it, vi } from "vitest";
import { act, screen } from "@testing-library/react";

// Added in QA pass for F-001: mocked Office.onReady gate (spec edge case: render nothing until ready).
describe("taskpane entry (mocked Office.onReady)", () => {
  it("renders nothing until Office.onReady resolves, then renders the panel (edge)", async () => {
    document.body.innerHTML = '<div id="container"></div>';
    let ready: (() => void) | undefined;
    (globalThis as any).Office = {
      onReady: vi.fn((cb: () => void) => {
        ready = cb;
      }),
      CoercionType: { Text: "text" },
      AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
      context: { mailbox: { item: undefined } },
    };
    vi.resetModules();
    await import("./index");
    expect(Office.onReady).toHaveBeenCalledTimes(1);
    expect(document.getElementById("container")?.innerHTML).toBe("");
    await act(async () => {
      ready?.();
    });
    expect(await screen.findByText("Open a message to use DIMSCALE Mail AI")).toBeTruthy();
  });
});
