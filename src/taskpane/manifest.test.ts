/* global process */

import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

// Added in QA pass for F-001 (AC-2): static guard on manifest.xml. A text check, not a schema check.
// Relies on the repo root as cwd, which `npm test` sets.
const manifest = readFileSync(resolve(process.cwd(), "manifest.xml"), "utf8");

describe("manifest.xml (AC-2)", () => {
  it("requests only the ReadItem permission", () => {
    expect([...manifest.matchAll(/<Permissions>([^<]*)<\/Permissions>/g)].map((m) => m[1])).toEqual(
      ["ReadItem"]
    );
    expect(manifest).not.toMatch(/ReadWriteItem|ReadWriteMailbox|Mail\.Send/);
  });

  it("activates only on the message read form", () => {
    expect([...manifest.matchAll(/<Rule xsi:type="ItemIs"[^>]*\/>/g)].map((m) => m[0])).toEqual([
      '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />',
    ]);
    expect(manifest).not.toMatch(
      /MessageComposeCommandSurface|AppointmentOrganizer|AppointmentAttendee/
    );
  });

  it("declares no Graph/SSO web application info and no send-related events", () => {
    expect(manifest).not.toMatch(/WebApplicationInfo|ItemSend|ItemInsert/);
  });
});
