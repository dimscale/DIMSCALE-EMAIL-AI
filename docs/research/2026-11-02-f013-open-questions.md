# F-013 open questions Q-C1, Q-C3, Q-C4: recommendations (2026-11-02)

- **Status:** research input for the user. The F-013 spec is **not** edited; the user answers and the spec is amended afterwards.
- **Spec:** [F-013](../specs/F-013-automatic-reply-drafts.md). Decision row: D17 in [status.md](../status.md).

## Facts checked (Microsoft Learn, read 2026-11-02)
| Fact | Source |
|---|---|
| Get message: delegated `Mail.ReadBasic` (least) or `Mail.Read`; application `Mail.ReadBasic.All` (least) or `Mail.Read`. Headers are returned with `$select=internetMessageHeaders` (documented example 2). | https://learn.microsoft.com/graph/api/message-get?view=graph-rest-1.0 |
| `internetMessageHeaders`: "The set includes message headers indicating the network path taken by a message from the sender to the recipient... Requires `$select` to retrieve. Read-only." `from` and `sender` differ only in sharing or delegation scenarios. `inferenceClassification`: "based on inferred relevance or importance, or on an explicit override" (`focused` or `other`). | https://learn.microsoft.com/graph/api/resources/message?view=graph-rest-1.0 |
| `Mail.ReadBasic`: all mail "except the body, previewBody, attachments, and any extended properties". Headers are **not named** in the exclusion list, but the pages read do not say they are returned under `Mail.ReadBasic`. **Unverified**, so not assumed. | F-013 spec table (permissions reference), https://learn.microsoft.com/graph/api/message-get?view=graph-rest-1.0 |
| List messages: delegated least `Mail.ReadBasic`; application least `Mail.ReadBasic.All`. The page does not say whether `internetMessageHeaders` can be selected in a list call (the resource page only says "requires `$select`"). **Unverified.** | https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0 |
| Exchange RBAC for Applications has a role **`Application Mail.ReadBasic`** (same limits as the Graph permission) besides `Application Mail.Read` and `Application Mail.ReadWrite`. Only `Mail.Read` is needed for subscriptions (F-013 table), so the app keeps `Mail.Read`; the extra role gives no benefit while the same app holds `Mail.Read`. | https://learn.microsoft.com/exchange/permissions-exo/application-rbac |
| Microsoft 365 stamps **Authentication-Results** on inbound messages with SPF, DKIM, DMARC and `compauth` results (`dmarc=pass\|fail\|bestguesspass\|none`, `compauth=pass\|fail\|softpass\|none`, with a `reason` code). "Despite a `compauth` failure, the message might still be allowed if other assessments don't indicate a suspicious nature." The page does not say what is stamped on mail that originates inside the tenant. | https://learn.microsoft.com/defender-office-365/message-headers-eop-mdo |
| Auto-responders should not answer messages with `Auto-Submitted` other than `no`, and may skip `List-*` and `Precedence: list\|junk\|bulk`. | https://www.rfc-editor.org/rfc/rfc3834 |
| Outlook contacts: delegated least `Contacts.Read`. People API is in maintenance mode. | https://learn.microsoft.com/graph/api/user-list-contacts?view=graph-rest-1.0, https://learn.microsoft.com/graph/api/user-list-people?view=graph-rest-1.0 |
| With Assignment required, any delegated permission needs admin consent. | https://learn.microsoft.com/entra/identity/enterprise-apps/assign-user-or-group-access-portal |

## Q-C1. Candidate source
**Recommendation: keep v1 as specified, recent Inbox senders (`Mail.ReadBasic`, 30 days) plus manual entry. Do not add `Contacts.Read` now.**

| Option | Pros | Cons |
|---|---|---|
| **Recent Inbox senders + manual entry (spec v1)** | One delegated scope, already in the MMSI request and reportedly consented on 2026-10-29. Shows only `from`. Works for the people who actually write to the user. | Only senders from the last 30 days. Includes newsletters and notifications (the no-reply filter removes only obvious ones). A rare contact is added by hand. |
| Add Outlook contacts (`Contacts.Read`) | Curated list, includes people who have not written recently. | A second delegated scope: another admin consent (Assignment required), another request to MMSI, a wider read (all contacts including phone numbers, unless `$select` limits it). Many users keep no contacts. |
| People API | Relevance-ranked. | Maintenance mode (rejected in the spec). |
| Sent Items recipients | Reflects who the user answers. | Needs a Sent Items read in the pane (`toRecipients` under `Mail.ReadBasic` is **unverified**, see F-016 Unverified 2). Not recommended for v1. |

Reasoning: the first option adds no scope and no admin work, and the manual field covers the gap. Revisit `Contacts.Read` only if the user reports that the list is not useful after live use. No spec change is needed.

## Q-C3. Filters for listed contacts
**Recommendation: keep filter 6 (only CC'd), turn filter 12 (Focused "Other") off by default. Keep every other filter. No other change.**

| # | Filter | For an explicitly listed contact | Reason |
|---|---|---|---|
| 3, 4, 5 | not inbox message, event, from self | keep | Structural, not about relevance. |
| **6** | owner not in `toRecipients` | **keep** | A listed contact who copies the user is usually informing, not asking. A draft would be noise. Trade-off: a real question sent with the user in CC is missed. Option: make it configurable (`REQUIRE_DIRECT_RECIPIENT`), as filter 8, 12 and 16 already are (F-013/AC-13). |
| 7, 8, 9, 10 | auto-submitted, auto-response suppress, bulk precedence, mailing list | keep | Machine mail from a listed address (out-of-office, list traffic) should not get a reply (RFC 3834). A human writing through a list is skipped; the user can answer by hand. |
| 11 | no-reply local part | keep | Also reject such an address at manual entry (F-013/AC-34 already rejects invalid ones; add this case when the spec is amended). |
| **12** | `inferenceClassification = other` | **turn off** | Microsoft defines it as inferred relevance or an explicit override. The user's list is an explicit signal; a listed contact landing in "Other" should still be answered. Trade-off: if the user's Focused Inbox is untrained, a listed sender may land in Other and still get a draft, which is the intended behaviour. |
| 13 | older than 24 h | keep | Bounds catch-up. |
| 14 | empty body | keep | Nothing to answer. |
| 15, 16 | domain skip, internal | keep defaults (empty; internal allowed) | Internal listed senders (for example the test contact) are answered. |
| 17 | daily caps | keep (30 per user, 200 global) | Cost bound. |

Spec impact if accepted: filter 12 default `on` becomes `off` (a configuration value, AC-12 and AC-13 test cases change for that one default); filter 6 unchanged.

## Q-C4. Forged `From`, `Authentication-Results` hardening
**What is available.** Headers need `$select=internetMessageHeaders` on a single Get message call. F-013 already makes that second read (filters 7 to 10 use it) with application `Mail.Read` under the RBAC role, so the header check costs **no extra Graph call**. Whether the same header read works with only `Mail.ReadBasic` (delegated or `.All`) is **not documented**: do not design around it. The pane never needs headers.

**What a header check can and cannot do.**
- `Authentication-Results` carries SPF, DKIM, DMARC and `compauth` for inbound mail. A forged external `From` normally shows `dmarc=fail` or `compauth=fail`, depending on the sender domain's published records (not verified by test).
- Microsoft says a failing message may still be delivered. A header check is therefore a second line, not a guarantee.
- It says nothing documented about mail from inside the tenant, so a missing header must not be treated as a failure.
- `from` versus `sender` is **not** an anti-forgery check (they differ only for delegation and shared mailboxes).
- Forwarded mail and some lists break SPF or DKIM; skipping on SPF or DKIM alone would cause false skips.

| Option | Pros | Cons |
|---|---|---|
| A. Accept in v1 (spec today) | No work. Impact is an unsent draft plus an OpenAI call and cost. | The forged text reaches OpenAI. Caps and the marker bound it only. |
| **B. Add filter 18 `auth_failed`: skip when `Authentication-Results` contains an explicit `dmarc=fail` or `compauth=fail`; a missing header, `none`, `bestguesspass`, `softpass`, SPF or DKIM alone do not skip.** | Free (headers already read). Blocks the common forged external sender before any OpenAI call. Low false-skip risk. | Parsing a header whose exact form with several `Authentication-Results` lines is undocumented: the parser must be tolerant and tested with real samples. Does not stop a spoof that passes DMARC. |
| C. Strict: skip unless `dmarc=pass` or `compauth=pass` | Strongest. | Skips internal mail (header contents unverified), forwarded mail and any domain without DMARC. Likely to silence the feature for real contacts. |

**Recommendation: B, as a small amendment (new filter 18, reason code `auth_failed`, default on, configurable), with the risk R10 kept as residual.** If the user prefers speed, take A now and add B after the first live week. C is not recommended.

Proposed ACs if B is accepted (to be numbered by the spec owner):
- Given an allow-listed fixture with `Authentication-Results: ... dmarc=fail ...`, the row ends `skipped` / `auth_failed`, with no OpenAI call and no `createReply`. Mocks.
- The same for `compauth=fail`. Mocks.
- An allow-listed fixture with no `Authentication-Results`, or with `dmarc=bestguesspass`, or `spf=fail` alone, reaches the OpenAI call. Mocks.
- Live: record the `Authentication-Results` of one real external test mail and one internal test mail from the pilot mailbox (headers only, no content) and whether a forged external sender, if the user can send one, ends `auth_failed`. Live (user).

## Decisions for the user
1. Q-C1: keep recent senders plus manual entry for v1 (recommended), or add `Contacts.Read` now?
2. Q-C3: keep filter 6 and turn filter 12 off (recommended)? Or make filter 6 configurable too?
3. Q-C4: add the `auth_failed` filter (B, recommended), accept the risk for v1 (A), or strict (C)?

## Not verified
- Whether `internetMessageHeaders` is returned under `Mail.ReadBasic` and in list calls.
- What `Authentication-Results` contains for mail sent from inside the tenant, and whether several such headers appear.
- Whether a forged-sender test mail can be produced inside this tenant (and whether Microsoft 365 spoof policy would move it to Junk or quarantine before it reaches the Inbox).
