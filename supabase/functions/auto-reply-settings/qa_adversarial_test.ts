// QA (adversarial) for F-013/AC-37 and AC-49: PUT contact validation with no-reply variants. No DB, no network.
// The "KNOWN LIMIT" test pins the Cyrillic look-alike (not handled), so a change is noticed.
import { assert, assertEquals, assertFalse } from "@std/assert";
import { parsePut } from "./handler.ts";

const body = (contacts: unknown[]) => ({ enabled: true, ownAddress: "owner@dimscale.com", contacts });

Deno.test("ADV F-013/AC-49: every no-reply variant is rejected, alone or next to a valid contact (whole PUT is null)", () => {
  const variants = [
    "noreply@x.com", " NoReply@X.com ", "NO-REPLY@x.com", "donotreply@x.com", "Do-Not-Reply@x.com",
    "mailer-daemon@x.com", "postmaster@x.com", "NOREPLY@X.COM\n", "\tpostmaster@x.com",
    "NoReply <noreply@x.com>", // display-name form: not a valid address
  ];
  for (const v of variants) {
    assertEquals(parsePut(body(["ana@cliente.pt", v])), null, JSON.stringify(v));
    assertEquals(parsePut(body([v])), null, JSON.stringify(v));
  }
});

Deno.test("ADV F-013/AC-49: plus-addressing, fullwidth and zero-width variants are rejected as no-reply", () => {
  for (const v of ["noreply+x@x.com", "ｎｏｒｅｐｌｙ@x.com", "noreply​@x.com"]) assertEquals(parsePut(body([v])), null, JSON.stringify(v));
});

Deno.test("ADV F-013/AC-49 KNOWN LIMIT: a Cyrillic look-alike is accepted; noreply.team and x-noreply stay accepted; pinned", () => {
  for (const v of ["noreply.team@x.com", "x-noreply@x.com", "nоreply@x.com" /* Cyrillic o */]) {
    const r = parsePut(body([v]));
    assert(r !== null, `accept: ${JSON.stringify(v)}`);
    assertEquals(r!.contacts, [v.trim().toLowerCase()]);
  }
});

Deno.test("ADV F-013/AC-37: duplicates after normalisation and the owner address are rejected (whole PUT null)", () => {
  assertEquals(parsePut(body(["ana@x.com", " ANA@x.com "])), null, "trim+case duplicate is 422 (spec: no duplicates)");
  assertEquals(parsePut(body(["ana@x.com", "noreply@x.com", "ANA@x.com"])), null, "duplicate next to a no-reply");
  assertEquals(parsePut(body(["OWNER@dimscale.com"])), null, "own address in any case");
  assertFalse(parsePut(body(["ana@x.com", "rui@x.com"])) === null, "two distinct valid contacts pass");
});
