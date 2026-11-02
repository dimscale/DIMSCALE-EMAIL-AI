import { assert, assertFalse } from "@std/assert";
import { hasLink, introducesNewLinks, parseOutput } from "./ai.ts";

const MUST_FLAG = [
  "http://evil.example", "HTTPS://X.y", "www.evil.example", "go //evil.example/x", "boss@evil.example",
  "evil.com", "visite evil.com.", "www.evil.com/x", "evil.com/path", "mail.evil.co.uk", "EVIL.COM", "a-b.c-d.io",
  "report.pt", // file name ending in a TLD: flagged on purpose (ambiguous with a domain)
  "http://1.2.3.4", "ip 192.168.0.1 aqui", "evil。com", "evil．com", "evil․com", "evil\u200B.com", "evil.\u200Bcom", "ev\u200Bil.co\u200B.uk", "evil\uFF0Ecom/x",
];
const MUST_PASS = [
  "Pagar a fatura até sexta.", "a/b e c/d", "Olá @Ana", "2026-10-29", "versão v1.2", "reunião às 9.30", "valor 1.234,56 EUR", "3.14",
  "etc. e tal", "ex.: isto", "Sr.Silva", "Bom dia.Obrigado", "Fim.Seguimos", "MyEdenred", "Edenred", "Obrigado.", "Confirmo o email da MyEdenred. Obrigado.", "index.html", "Dr. Costa, e.g. amanhã", "1.2.3",
];

Deno.test("hasLink: table of must-flag strings", () => {
  for (const s of MUST_FLAG) assert(hasLink(s), JSON.stringify(s));
});
Deno.test("hasLink: table of must-pass strings", () => {
  for (const s of MUST_PASS) assertFalse(hasLink(s), JSON.stringify(s));
});
Deno.test("introducesNewLinks: bare domain new vs already in the email", () => {
  assert(introducesNewLinks("Veja evil.com", "Olá"));
  assert(introducesNewLinks("Veja evil\u200B.com", "Olá"));
  assertFalse(introducesNewLinks("Veja exemplo.pt", "Contacte exemplo.pt"));
  assertFalse(introducesNewLinks("Bom dia.Obrigado", "Olá"));
});

Deno.test("hasLink: invisible characters inside a domain are stripped and flagged", () => {
  for (const c of ["\u2066", "\u2069", "\u180E", "\u034F", "\uFE0F", "\u{E0041}", "\u061C", "\u3164"]) {
    assert(hasLink(`evil${c}.com`), `U+${c.codePointAt(0)!.toString(16)} before dot`);
    assert(hasLink(`ev${c}il.com`), `U+${c.codePointAt(0)!.toString(16)} inside label`);
  }
});
Deno.test("hasLink: normal accented text is untouched, new TLDs and punycode flagged", () => {
  for (const s of ["coração, avião e café", "Atenção à reunião"]) assertFalse(hasLink(s), s);
  for (const s of ["a.click", "x.icu", "promo.work", "site.bank", "evil.xn--p1ai", "xn--80ak6aa92e.com"]) assert(hasLink(s), s);
});
Deno.test("introducesNewLinks: token compare, not substring", () => {
  assert(introducesNewLinks("evil.co", "veja evil.com"));
  assert(introducesNewLinks("l.com", "veja evil.com"));
  assert(introducesNewLinks("1.2.3.4", "ip 11.2.3.45"));
  assert(introducesNewLinks("evil.com/steal", "veja evil.com"));
  assert(introducesNewLinks("ana@exemplo.pt", "veja b.ana@exemplo.pt"));
  assertFalse(introducesNewLinks("evil.com/steal", "veja evil.com/steal"));
  assertFalse(introducesNewLinks("evil.com/steal", "veja https://www.evil.com/steal."));
  assertFalse(introducesNewLinks("evil.com", "veja evil.com/steal"));
  assertFalse(introducesNewLinks("mail.evil.com", "veja evil.com"));
  assertFalse(introducesNewLinks("evil.com", "veja mail.evil.com"));
  assertFalse(introducesNewLinks("ana@exemplo.pt", "Escreva a Ana@Exemplo.pt, obrigado"));
  assertFalse(introducesNewLinks("evil.com", "veja evil\u200B.com"));
  assertFalse(introducesNewLinks("ip 1.2.3.4", "ip 1.2.3.4"));
});
Deno.test("hasLink: 16,000 chars stays fast", () => {
  const t = performance.now();
  hasLink("a.b.c.d ".repeat(2000));
  introducesNewLinks("Olá ".repeat(4000), "a-b.c-d ".repeat(2000));
  assert(performance.now() - t < 1000);
});
Deno.test("parseOutput: a reply that repeats the sender name MyEdenred (no dot) is accepted", () => {
  const email = "Sender: MyEdenred appmyedenred.pt";
  const r = parseOutput({ status: "completed", output: [{ type: "message", content: [{ type: "output_text", text: JSON.stringify({ reply: "Olá, obrigado. Cumprimentos, MyEdenred." }) }] }] }, email);
  assert(r.ok);
});
