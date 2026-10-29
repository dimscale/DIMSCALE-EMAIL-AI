/* global globalThis */
import { afterEach, describe, expect, it, vi } from "vitest";
import { openReplyForm } from "./replyForm";

const office = (supported: boolean) => {
  (globalThis as any).Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: { requirements: { isSetSupported: vi.fn(() => supported) } },
  };
};
afterEach(() => delete (globalThis as any).Office);

describe("openReplyForm (F-003/AC-24)", () => {
  it("uses displayReplyFormAsync on Mailbox 1.9 and reports the callback status", async () => {
    office(true);
    const ok = {
      displayReplyFormAsync: vi.fn((_h: string, cb: any) => cb({ status: "succeeded" })),
      displayReplyForm: vi.fn(),
    };
    expect(await openReplyForm(ok as any, "<h>")).toBe(true);
    expect(ok.displayReplyFormAsync).toHaveBeenCalledTimes(1);
    expect(ok.displayReplyFormAsync.mock.calls[0][0]).toBe("<h>");
    expect(ok.displayReplyForm).not.toHaveBeenCalled();
    expect((globalThis as any).Office.context.requirements.isSetSupported).toHaveBeenCalledWith(
      "Mailbox",
      "1.9"
    );
    const bad = { displayReplyFormAsync: vi.fn((_h: string, cb: any) => cb({ status: "failed" })) };
    expect(await openReplyForm(bad as any, "x")).toBe(false);
  });
  it("falls back to displayReplyForm below 1.9", async () => {
    office(false);
    const it = { displayReplyFormAsync: vi.fn(), displayReplyForm: vi.fn() };
    expect(await openReplyForm(it as any, "<h>")).toBe(true);
    expect(it.displayReplyForm).toHaveBeenCalledWith("<h>");
    expect(it.displayReplyFormAsync).not.toHaveBeenCalled();
  });
  it("returns false on a thrown exception or no item", async () => {
    office(true);
    const boom = {
      displayReplyFormAsync: () => {
        throw new Error("32KB");
      },
    };
    expect(await openReplyForm(boom as any, "x")).toBe(false);
    expect(await openReplyForm(null, "x")).toBe(false);
  });
});
