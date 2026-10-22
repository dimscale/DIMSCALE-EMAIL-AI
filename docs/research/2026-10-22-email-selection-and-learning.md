# Research: choosing which emails to answer, and "learning" from chosen emails

- **Date:** 2026-10-22
- **Author:** requirements-architect agent
- **Status:** Research only. Not a spec, not approved. It changes no existing file and no code.
- **Requests (user, 2026-10-22, translated):**
  1. "A selection box (checkboxes) for which emails the AI should read and **learn from**, to improve."
  2. "A box listing the emails where I can **select which ones to reply to**", instead of, or as well as, fully automatic drafts for every incoming email.
- **Inputs read:** `CLAUDE.md`, `docs/status.md`, `docs/architecture/overview.md`, `docs/architecture/security.md`, `docs/product/requirements.md`, `docs/specs/F-002-authenticated-backend.md`, `docs/specs/F-012-taskpane-design.md`, `manifest.xml`.
- **Inputs not found:** `docs/decisions/0005-server-side-mail-processing.md` and `docs/specs/F-013-automatic-reply-drafts.md` are **not on disk** in this working tree. `status.md` says both are "being written". Everything below about ADR-0005 and F-013 is based on the orchestrator's summary: server-side automatic drafts using Graph change notifications and Supabase, app-only access with Exchange RBAC for Applications scoped to a group, and never `Mail.Send`. Re-check this section against the real files once they exist.

---

## 1. Summary and recommendation

**Short version**
- **Request 2 (choose which emails to answer).** Build a **"manual" mode first**: a list of recent Inbox emails inside the pane, with checkboxes. The user ticks some emails and presses "Criar rascunhos". For each ticked email the pane:
  1. reads it through Microsoft Graph with the user's own (delegated) permission;
  2. sends the minimized text to the DIMSCALE backend for the AI suggestion;
  3. creates a **reply draft** in the user's Drafts folder.

  Nothing is sent. This is the backlog's F-008 (list) plus F-009 (Graph draft), with checkboxes added.
  - It needs **no manifest change and no admin re-upload**.
  - It needs **no app-only permission, no change notifications and no stored email**, because the user is present.
  - So the heavy parts of ADR-0005 (app-only access, Exchange RBAC for Applications, webhooks) are needed **only for the later "automatic" mode**.
- **Do not use Outlook's own multi-select (Ctrl+click) for v1.** Microsoft documents that `getSelectedItemsAsync` and `loadItemByIdAsync` need the **read/write mailbox** add-in permission. That is the highest level: it includes "sending items", and such an add-in needs an administrator to install it. That is more privilege than the Graph route needs. The body-reading part (Mailbox 1.15) is also not available on Mac.
- **Request 1 ("learn").** In v1, store a per-user **style profile**, not emails:
  - The user ticks some of their **own sent replies**.
  - The backend sends only the new part of each reply (Graph `uniqueBody`, plain text, truncated) to the AI once.
  - The AI returns a short description of tone, greeting, sign-off and length. The user reviews and edits it.
  - Only that text is saved in Supabase, keyed by the user's Entra object ID, and it can be deleted with one button.
- **Not recommended:**
  - **Fine-tuning** is off the table: OpenAI says it is "winding down the fine-tuning platform" and that it "is no longer accessible to new users".
  - **Few-shot retrieval** (storing real emails plus embeddings in Postgres with pgvector) works, but it stores third-party email content. It should wait until the DPO has assessed it.
- **"Learning which emails deserve a reply":** use user-confirmed rules (for example "always ignore this sender or domain") plus Outlook's own Focused/Other flag. No machine learning in v1.
- **Blockers that still apply:** the OpenAI and GDPR decision (D5), plus Entra consent for the delegated Graph scopes (admin).

**Recommended v1 scope (proposal)**

| Need | v1 proposal | Permission | Manifest change | Admin action |
|---|---|---|---|---|
| List emails with checkboxes | Graph `GET /me/mailFolders/inbox/messages` with `$select` (no body) | Delegated `Mail.ReadBasic` | None | Consent for the scope |
| Draft replies for the ticked emails | Graph read of the body, then backend AI, then `createReply` + `PATCH` | Delegated `Mail.ReadWrite` (does not include send) | None | Consent for the scope |
| "Learn my style" | Style profile built from ticked sent replies; only the profile is stored | Delegated `Mail.Read` (or `Mail.ReadWrite` once granted) | None | Same consent |
| "Learn which to answer" | User-confirmed ignore/always rules plus Focused/Other | Same as the list | None | None |
| Automatic drafts with Outlook closed | Later (F-013, ADR-0005) | App-only, RBAC-scoped | None | Yes, plus the DPO |

---

## 2. Where each capability sits

| Capability | Tier | Mechanism |
|---|---|---|
| Act on the open email (summary, suggested reply, open a reply form) | Active item | Office.js, `ReadItem` (today) |
| Act on several emails selected in Outlook's own list | Active item, extended | Office.js item multi-select: read/write mailbox permission, Mailbox 1.13 or 1.15 (section 3) |
| List emails in the pane and pick some | Mailbox | Graph, delegated (section 4) |
| Create reply drafts for picked emails | Mailbox | Graph `createReply`, delegated `Mail.ReadWrite` |
| Act while Outlook is closed | Backend + Mailbox | Graph app-only plus change notifications (ADR-0005, F-013) |
| Store "learned" data | Backend | Supabase Postgres (new: the backend becomes stateful) |

`requirements.md` lists **"Automatic rules or filters, bulk actions"** as out of scope. Both requests are bulk actions, and F-013 is automatic. Adopting either needs the user to amend `requirements.md` (a product decision, open question U1).

---

## 3. Topic A: Outlook item multi-select (Office.js)

### Verified facts
| Fact | Source |
|---|---|
| Multi-select was introduced in **Mailbox 1.13**. `loadItemByIdAsync` and `unloadAsync` were introduced in **Mailbox 1.15**. | [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| XML manifest: add the `SupportsMultiSelect` child element, set to `true`, to the `<Action xsi:type="ShowTaskpane">`. The `ExtensionPoint` must be `MessageReadCommandSurface` or `MessageComposeCommandSurface`. The element is valid only in the Mail 1.1 `VersionOverrides` schema and is associated with Mailbox 1.13. The tutorial also sets `<Permissions>ReadWriteMailbox</Permissions>`, `SupportsPinning`, `SupportsNoItemContext` and `DefaultMinVersion="1.15"`. | [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select), [Action element](https://learn.microsoft.com/javascript/api/manifest/action) |
| "If you turn on the item multi-select feature in your add-in, your add-in will automatically support the no item context feature, even if it isn't explicitly configured in the manifest." | [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| `Office.context.mailbox.getSelectedItemsAsync`: Mailbox 1.13, **minimum permission level: read/write mailbox**, Compose or Read mode. "An add-in can activate on a maximum of 100 messages at a time." It returns, per message: attachment flag, conversation ID, internet message ID, item ID (EWS format), item mode, item type (`Message` only) and subject. It returns **no sender and no body**. | [Office.Mailbox (1.15)](https://learn.microsoft.com/javascript/api/outlook/office.mailbox?view=outlook-js-1.15), [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| `SelectedItemsChanged` event: registered with `Office.context.mailbox.addHandlerAsync` (Mailbox 1.5, read item). The event itself is Mailbox 1.13. "Events are only available in task pane add-ins." | [Office.Mailbox (1.15)](https://learn.microsoft.com/javascript/api/outlook/office.mailbox?view=outlook-js-1.15) |
| `Office.context.mailbox.loadItemByIdAsync(itemId)`: Mailbox 1.15, **minimum permission level: read/write mailbox**. "Only one mail item can be loaded at a time... you must call `unloadAsync` after processing the item." "Can only be called on messages in the same mailbox." It returns `LoadedMessageRead` or `LoadedMessageCompose`. Not supported in event-based activation. | [Office.Mailbox (1.15)](https://learn.microsoft.com/javascript/api/outlook/office.mailbox?view=outlook-js-1.15), [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| `LoadedMessageRead`: Mailbox 1.15, interface minimum permission read/write item, Message Read mode. It has a `body` property (read item); "Only the `getAsync` and `getTypeAsync` methods of the Body object are supported". It also lists `from`, `sender`, `to`, `cc`, `dateTimeCreated`, `categories`, `displayReplyFormAsync`, `displayReplyAllFormAsync` and `unloadAsync`. "Typically, you can only run get operations on a selected message that's loaded"; categories are the documented exception. | [LoadedMessageRead](https://learn.microsoft.com/javascript/api/outlook/office.loadedmessageread), [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| Limits: one Exchange mailbox at a time; one folder at a time (unless Conversations view); the add-in must have a task pane; the Reading Pane must be on (unless no-item-context applies); messages only. On Windows (new and classic) and the web, "the task pane doesn't stay open when the user switches to a new item" unless the user pins it. On Mac it stays open. | [Item multi-select](https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select) |
| Permission levels: **read/write mailbox** = everything in read/write item **plus** "creating, reading, writing items and folders", "**sending items**" and `makeEwsRequestAsync`. "Outlook add-ins that request the read/write mailbox permission require explicit administrator privilege to install." | [Outlook add-in permissions](https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions), [Privacy and security](https://learn.microsoft.com/office/dev/add-ins/concepts/privacy-and-security) |

### Client support (from the requirement-set tables)
| Client | Mailbox 1.13 (`getSelectedItemsAsync`, event) | Mailbox 1.15 (`loadItemByIdAsync`, body of selected items) |
|---|---|---|
| Classic Outlook on Windows, Microsoft 365 or retail perpetual | Version 2304 (Build 16327.20248) | **Version 2412 (Build 18324.20172)** |
| Classic Outlook, volume-licensed LTSC 2024 | Yes (up to 1.14) | **No** |
| Classic Outlook, LTSC 2021 | No (up to 1.9) | No |
| New Outlook on Windows, Outlook on the web | Yes | Yes |
| Outlook on Mac (new UI) | Yes (up to 1.14) | **No** |
| Android, iOS | No (1.5) | No |

Source: [Outlook requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). Every client that can sign in (NAA needs classic Version 2409 or later, F-002) has 1.13. Reading the bodies of selected items needs 1.15, so classic 2412 or later, and it is not available on Mac.

### Assessment
- **It can work without Graph:** the user Ctrl+clicks up to 100 emails in Outlook's list, the pane reads each one in turn (`loadItemByIdAsync`, then `body.getAsync`, then `unloadAsync`) and sends minimized text to the backend.
- **But it costs the highest add-in permission:**
  - `ReadWriteMailbox` (manifest `<Permissions>`) replaces `ReadItem`, and it includes "sending items".
  - The code would never send, but the permission is broader than needed and broader than the Graph route (`Mail.ReadWrite`: "Does not include permission to send mail").
  - It conflicts with CLAUDE.md rule 7 (least privilege; active-item features use `ReadItem`) and with the spirit of rule 1.
  - It also requires administrator installation. The production manifest is admin-deployed anyway (F-011), so this is a governance cost, not a blocker.
- **Drafts are still not possible** with Office.js alone for non-active items. `LoadedMessageRead` lists `displayReplyFormAsync`, but its behavior on a loaded item is undocumented beyond the member list (unverified). Creating a draft in Drafts would still need Graph `Mail.ReadWrite`.
- **Manifest change:** `Permissions`, `SupportsMultiSelect`, `SupportsPinning`, and possibly `DefaultMinVersion`. That means a `<Version>` bump and an **admin re-upload** (F-011 "Manifest changed later").
- **The pane must handle new states:** no item (no-item-context becomes automatic), several items, and `SelectedItemsChanged` while it is open. F-001 and F-012 do not handle `ItemChanged` today.
- **Verdict:** not recommended for v1. Keep it as an option if the user strongly prefers Ctrl+click in Outlook's own list over a list in the pane (open question U3).

### Unverified for A
- The element name's casing. The Action reference heading says `SupportsMultiselect`, while the XML snippets on the same page and in the tutorial use `SupportsMultiSelect`. Confirm with `npm run validate` before any manifest change.
- Whether `SupportsMultiSelect` works with `DefaultMinVersion` left at 1.3 (the tutorial raises it to 1.15; the Action page only says "associated with" 1.13).
- Whether `loadItemByIdAsync` depends on the EWS endpoint that Exchange Online is disabling. Microsoft says "October 2026: EWS starts to be disabled globally" and "April 2027: EWS is fully disabled" ([EWS deprecation](https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online)). The Office.js pages only say the method uses EWS item IDs; they say nothing about the service dependency.

---

## 4. Topic B: email list with checkboxes in the pane (Microsoft Graph, delegated)

### Verified facts
| Fact | Source |
|---|---|
| `GET /me/mailFolders/{id}/messages`. Well-known names such as `inbox`, `sentitems` and `drafts` can replace the folder ID "regardless of the locale". | [List messages](https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0), [mailFolder](https://learn.microsoft.com/graph/api/resources/mailfolder?view=graph-rest-1.0) |
| Least-privileged delegated permission for listing: **`Mail.ReadBasic`**. Higher: `Mail.ReadWrite`, `Mail.Read`. | [List messages](https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0) |
| `Mail.ReadBasic` (delegated): "read email in the signed-in user's mailbox except body, previewBody, attachments and any extended properties." `Mail.Read`: "read the signed-in user's mailbox." `Mail.ReadWrite`: "create, read, update, and delete email in user mailboxes. **Does not include permission to send mail.**" Admin consent required, delegated: **No** for all three. | [Permissions reference](https://learn.microsoft.com/graph/permissions-reference) |
| Default page size 10. `$top` from 1 to 1000. Use `$select` to reduce payload. Follow `@odata.nextLink` without editing `$skip`. With `$filter` plus `$orderby`, the ordered properties must also appear in the filter, first and in the same order; otherwise the error is `InefficientFilter`. | [List messages](https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0) |
| `Prefer: outlook.body-content-type` can be "text" or "html". The same page also says "Currently, this operation returns message bodies in only HTML format" (contradictory, see Unverified). | [List messages](https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0) |
| `uniqueBody`: "The part of the body of the message that is unique to the current message", not returned by default, request it with `$select=uniqueBody`. `bodyPreview`: first 255 characters, text. `inferenceClassification`: `focused` or `other`. `categories` is a message property. `id` changes when the item moves, unless `Prefer: IdType="ImmutableId"` is used. | [message resource](https://learn.microsoft.com/graph/api/resources/message?view=graph-rest-1.0) |
| `createReply`: "Create a draft to reply"; "Send the draft message in a subsequent operation" (never called by us). Least privileged: `Mail.ReadWrite` (delegated and application). In JSON, specify either a `comment` or `body`, not both. | [createReply](https://learn.microsoft.com/graph/api/message-createreply?view=graph-rest-1.0) |
| NAA "supports incremental and dynamic consent so always request the minimum scopes needed". The pattern is `acquireTokenSilent`, with `acquireTokenPopup` on `InteractionRequiredAuthError`. | [Enable NAA](https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in) |
| Hand-off to Outlook's own form: `Office.context.mailbox.displayMessageFormAsync(itemId)` (Mailbox 1.9, read item) opens an existing message from its EWS ID. `convertToEwsId` (Mailbox 1.3, **restricted**) converts a Graph/REST ID to EWS, and `convertToRestId` does the reverse. Neither is supported on Android or iOS. | [Office.Mailbox (1.15)](https://learn.microsoft.com/javascript/api/outlook/office.mailbox?view=outlook-js-1.15) |

### Proposed calls (for a future spec; not code)
1. **List (no body, `Mail.ReadBasic`):**
   `GET /me/mailFolders/inbox/messages?$select=id,subject,from,receivedDateTime,isRead,hasAttachments,inferenceClassification,conversationId&$filter=receivedDateTime ge {start}&$orderby=receivedDateTime desc&$top=25`
   This complies with the `$filter`/`$orderby` rule. No preview text is shown, because `Mail.ReadBasic` excludes it. That is more private, at a small usability cost (open question U5).
2. **Read a ticked email (`Mail.Read`, or `Mail.ReadWrite` if already granted):**
   `GET /me/messages/{id}?$select=subject,from,uniqueBody` with `Prefer: outlook.body-content-type="text"`, truncated to the spec limit before it goes to the backend.
3. **Create the draft (`Mail.ReadWrite`):** `POST /me/messages/{id}/createReply` with the AI text as an escaped `comment` (or `PATCH` the draft body afterwards). Never `/send`.
4. **"Open in Outlook" per row (no new permission):** `convertToEwsId`, then `displayMessageFormAsync`, so the user can open the email and use the existing active-item tools.

### How this fits the current architecture
- It fits the existing rule that **the backend never receives Graph tokens** (security.md):
  1. the pane reads the email through Graph;
  2. it sends minimized text to the Supabase function with the `access_as_user` token;
  3. it gets the suggestion back;
  4. it creates the draft itself through Graph.
- It works wherever NAA works (classic 2409 or later, web, new Outlook, Mac 16.89 or later per F-002), so it is **broader than A** (it includes Mac).
- Selection state stays in memory only (no `localStorage`: the Pages origin is shared, F-011).
- **Consent:** the app registration gets the delegated Graph scopes added in Entra, requested at runtime one by one (`Mail.ReadBasic` when the list opens, `Mail.ReadWrite` on the first "Criar rascunhos"). Delegated admin consent is "No" per Microsoft, but the tenant admin is already granting consent for `access_as_user` (status.md), so the same admin step applies. **No manifest change.**
- **Note on the pending revocation.** status.md asks the admin to revoke "the old delegated Graph grants (`Mail.ReadWrite`, `offline_access`, `User.Read`)". This proposal would re-introduce `Mail.ReadWrite` on purpose, for a documented feature. The user should tell the admin this explicitly so the two requests are not seen as contradictory (open question A1).

### Unverified for B
- Whether `GET /me/messages/{id}` honors `Prefer: outlook.body-content-type="text"` for `uniqueBody`. The list page both documents the header and says bodies are "only HTML". The get-message page was not fetched. Test it live; the fallback is to strip HTML in the pane before sending.
- Whether `Mail.ReadBasic` returns `inferenceClassification`, `isRead` and `hasAttachments`. Inferred yes from the "except body, previewBody, attachments and extended properties" wording, not stated per property.
- A reliable "already replied to" signal for the list. Not researched. The list would show all recent emails, and the user decides.

---

## 5. Comparison: A (Outlook multi-select) vs B (pane list via Graph) vs a category trigger

A third option came up during research: the user marks emails in Outlook itself with a **category** (for example "IA – Responder"). Categories already support Outlook's native multi-select on every client, mobile included. The F-013 server reacts to the change. Microsoft says that without a `$filter`, "changing any other property of a message in that folder would result in an `Updated` notification", and that message subscriptions accept `Mail.ReadBasic` or `Mail.Read` ([Outlook change notifications](https://learn.microsoft.com/graph/outlook-change-notifications-overview)). That a category change counts as such an update is **inferred**. A subscription `$filter` on `categories` is shown only in Microsoft Q&A threads, not in reference docs (**unverified**).

| | A: Office.js multi-select | B: pane list via Graph (recommended) | C: Outlook category + F-013 server |
|---|---|---|---|
| Where the user selects | Outlook's message list (Ctrl+click) | Checkbox list in the pane | Outlook's list, by assigning a category |
| Add-in manifest permission | **ReadWriteMailbox** (includes "sending items") | ReadItem (unchanged) | ReadItem (unchanged) |
| Graph permission | None to read; `Mail.ReadWrite` for drafts | `Mail.ReadBasic` (list), `Mail.Read`/`Mail.ReadWrite` (body, draft), delegated | App-only (F-013): read plus `Mail.ReadWrite`, RBAC-scoped |
| Manifest change and admin re-upload | Yes; admin-only install | No | No |
| Entra or Exchange admin work | Consent for drafts only | Delegated consent | App-only consent plus RBAC for Applications plus webhook |
| Clients | 1.13 select: Windows 2304+, web, new, Mac. 1.15 body: Windows 2412+, web, new; **not Mac** | Every NAA client (Windows 2409+, web, new, Mac 16.89+) | Any client, mobile included (the category is set in Outlook) |
| Max per action | 100 (platform) | Our cap (proposal 10–20) | Unlimited (server-paced) |
| Works with Outlook closed | No | No | Yes (drafts are ready later) |
| Data stored by us | None | None | Per ADR-0005 (processing state) |
| Complexity | Medium (sequential load and unload, new pane states) | Medium (list UI, Graph client, consent states) | High (all of F-013) |
| Least privilege | Worst | Best | Medium (app-only, but scoped) |

---

## 6. Topic C: "learning" from selected emails

The model does not "learn" in any of these options except (iii). Options (i), (ii) and (iv) add **context to the prompt**. UI wording should say "Perfil de estilo" or "Exemplos", not "a IA aprende", so no capability is overpromised (F-012 principle 2).

### Verified OpenAI and Supabase facts used below
| Fact | Source |
|---|---|
| "Data sent to the OpenAI API is not used to train or improve OpenAI models unless you explicitly opt in." Abuse-monitoring logs are "retained for up to 30 days, unless longer retention is required by law". | [Your data (OpenAI)](https://developers.openai.com/api/docs/guides/your-data) |
| Application state: `/v1/chat/completions` none (exceptions: audio 1 h, prompt-cache tensors up to 24 h); `/v1/responses` at least 30 days when `store` is true (the default); `/v1/embeddings` none; `/v1/files` and `/v1/fine_tuning/jobs` "until deleted". Zero Data Retention and Modified Abuse Monitoring need OpenAI approval; files and fine-tuning are not ZDR-eligible. | same |
| EU data residency (`eu.api.openai.com`) "Requires Zero Data Retention, Modified Abuse Monitoring, Private Retention with PSP, or Safety Retention". It covers chat completions, responses and embeddings among others. | same |
| Fine-tuning: "OpenAI is winding down the fine-tuning platform." "The platform is no longer accessible to new users." Existing users can create jobs "until January 6, 2027, unless the base model shuts down earlier". Supported bases are the gpt-4.1 family: `gpt-4.1-nano` shuts down October 23, 2026, and the others January 14, 2027. Minimum 10 examples; "We recommend starting with 50 well-crafted demonstrations". | [Model optimization](https://developers.openai.com/api/docs/guides/model-optimization), [Supervised fine-tuning](https://developers.openai.com/api/docs/guides/supervised-fine-tuning) |
| OpenAI's optimization guidance: evals first, then prompting with few-shot examples. "The prompt engineering process may be all you need." | [Model optimization](https://developers.openai.com/api/docs/guides/model-optimization) |
| Embeddings: `text-embedding-3-small` (1536 dimensions) and `text-embedding-3-large` (3072). The `dimensions` parameter shortens vectors. Max input 8192 tokens. Priced per input token (illustration: about 62,500 pages per dollar for `-3-small`). Retrieval is ranked by cosine similarity. | [Embeddings](https://developers.openai.com/api/docs/guides/embeddings) |
| Supabase pgvector: `create extension vector with schema extensions;`, column `extensions.vector(n)`, HNSW and IVFFlat indexes exist. | [pgvector](https://supabase.com/docs/guides/database/extensions/pgvector) |
| Supabase Vault stores "encrypted secrets and other data"; the encryption key "is never stored in the database alongside the encrypted data". There is no guidance for bulk row data. | [Vault](https://supabase.com/docs/guides/database/vault) |
| "Supabase projects are encrypted at rest by default." pgsodium "will be deprecated. Use Supabase Vault instead." Transparent column encryption is discouraged. | [pgsodium](https://supabase.com/docs/guides/database/extensions/pgsodium) |
| Backups: Pro has the last 7 days of daily backups, Team 14, Enterprise up to 30. PITR options are 7, 14 or 28 days. "When you delete a project, we permanently remove all associated data, including any backups." Whether deleted rows persist in backups until they expire, and the backup region, **are not stated**. | [Backups](https://supabase.com/docs/guides/platform/backups) |

### (i) Per-user style profile (recommended for v1)
- **Flow:**
  1. In the pane, the user ticks 5 to 20 of their own replies from Sent Items. The list uses `Mail.ReadBasic` on `sentitems`.
  2. The pane reads each `uniqueBody` with Graph (`Mail.Read`), so only the user's own new text is used, not the quoted thread. It truncates each one (proposal: 1,500 characters).
  3. The pane sends them in one request to a new backend operation.
  4. The backend asks the model for a compact profile with a fixed schema: language, formality, greeting pattern, sign-off pattern, typical length, recurring phrases, things to avoid, about 800 characters at most. The output is validated, and anything that looks like an address or phone number is rejected.
  5. The pane shows the profile. The user edits and saves it. Later reply prompts include the profile as **system-side style guidance**, never as instructions taken from email.
  - A no-new-permission variant: the user opens one sent email and clicks "Usar como exemplo" (active item, `ReadItem`). This is slower, one at a time. Whether the add-in activates on items in Sent Items is expected but must be checked live (**unverified**).
- **Leaves Microsoft 365:** the selected replies' text, once per build, to OpenAI through the backend. It can include recipients' names in greetings and third-party details the user wrote.
- **Stored:** the profile text, the user's `oid`, and created/updated timestamps in Supabase (eu-west-1, D11). No email text, no message IDs, no addresses.
- **How long:** until the user deletes it, or until offboarding (process to be defined, open question P6). OpenAI keeps abuse logs up to 30 days (chat completions keeps no application state). If the Responses API is used, set `store: false`.
- **GDPR and DPO:** the profile is personal data about the user (derived), but low-volume and visible to them. The build step transiently processes third-party data at OpenAI, the same as every reply feature (D5). Needs a notice line and a lawful basis.
- **Delete on request:** a "Apagar perfil" button deletes the row by `oid`. An admin script deletes by `oid` for leavers. Backups keep it up to the plan's retention (7 days on Pro; **inferred**, not stated).
- **Cost and complexity:** one model call per build, plus about 200 extra prompt tokens per draft. Low to medium: one table, one new function operation, and the first database use by the backend (see section 9).

### (ii) Few-shot retrieval with stored examples and pgvector (later, only after a DPO assessment)
- **Flow:**
  1. The user ticks "good examples": pairs of an incoming email and their reply.
  2. The backend stores both texts (truncated) and an embedding of the incoming text (`text-embedding-3-small`, possibly shortened with `dimensions`).
  3. For each new email, it embeds the email, retrieves the top 2 or 3 similar examples by cosine similarity, and adds them to the prompt.
- **Leaves Microsoft 365:** the example texts (for embedding and again in every prompt that uses them), plus each new email for its embedding.
- **Stored:** **third-party email content** (senders' text and names) plus embeddings. Encrypted at rest by default. Field-level encryption would need custom application-level encryption: Vault has no bulk guidance and pgsodium is being deprecated. Embeddings derived from email should be treated as personal data (DPO to confirm).
- **How long:** must be fixed by the spec (proposal: at most 50 examples per user, each expiring after 180 days, deletable one by one). Backups as in (i).
- **GDPR and DPO:** materially higher risk. It creates a store of correspondence outside Microsoft 365, likely needs a DPIA, and needs a retention schedule and a subject-access/erasure path that covers the third parties in the emails, not only the user.
- **Delete:** per example, per user, and an admin purge by `oid`. Third-party erasure requests are hard: you have to search stored text for that person.
- **Cost and complexity:** embeddings are cheap, but every draft prompt grows by the examples. Medium to high.
- **Lower-risk alternative:** a user-written **"respostas modelo"** library (snippets the user types, such as standard pricing or availability answers). It reaches most of the benefit with no stored email content (open question U7).

### (iii) OpenAI fine-tuning (not recommended, and not available)
- OpenAI says the platform "is no longer accessible to new users". It is winding down, and the supported base models shut down between October 23, 2026 and January 14, 2027. DIMSCALE has no fine-tuning history in this project, so this is effectively unavailable.
- Even if it were available:
  - training files (`/v1/files`) and job state are kept "until deleted" and are not eligible for Zero Data Retention;
  - it needs at least 10 examples (50 recommended) of real email content;
  - the model cannot "unlearn" one person's data short of deleting the whole model.
- Verdict: excluded.

### (iv) Learning which emails deserve a reply
- **Signal:** the user's ticks in the list ("responder" or "ignorar"), or the category from option C.
- **(iv-a) User-confirmed rules (recommended for v1):**
  - When the user ignores the same sender or domain several times in one session, the pane offers "Ignorar sempre {domínio}?". Only a confirmed rule is stored: the pattern, the action and `oid`.
  - It also uses the free Focused/Other signal (`inferenceClassification`, readable with `Mail.ReadBasic`, **inferred**) as a pre-filter.
  - Leaves Microsoft 365: nothing new. Stored: rules (they contain third-party addresses or domains, which is small personal data). Delete: per rule, per user.
- **(iv-b) Few-shot classifier:** store labelled examples (subject, sender domain, label) and add them to a classification prompt. This sends subjects to OpenAI on every classification and stores subjects. Medium risk.
- **(iv-c) Embedding classifier:** like (ii), but on subjects. Same storage concerns.
- **Verdict:** v1 is (iv-a). (iv-b) is worth revisiting with F-013's "assisted" mode once real ticks show whether rules are enough.

### Comparison of the learning options
| | (i) Style profile | (ii) Few-shot plus pgvector | (iii) Fine-tuning | (iv-a) Rules | (iv-b) Few-shot classifier |
|---|---|---|---|---|---|
| Improves | Tone and format of drafts | Content of drafts | Both | Which emails get drafts | Which emails get drafts |
| Email content stored by DIMSCALE | **No** (derived text only) | **Yes** | At OpenAI (files, model) | No (addresses or domains only) | Subjects |
| Extra data sent to OpenAI | Once per build | Every draft (examples) plus embeddings | Training set | None | Every classification |
| Erasure | One row | Rows plus third-party search | Delete model and files | One row per rule | Rows |
| DPO effort | Low | High (likely DPIA) | n/a | Low | Medium |
| Complexity | Low–medium | Medium–high | n/a | Low | Medium |
| Available | Yes | Yes | **No (closed to new users)** | Yes | Yes |
| Recommended | **v1** | Later, if (i) plus templates fall short | No | **v1** | Later, with "assisted" |

---

## 7. Topic D: integration with F-013 (modes)

| Mode | What the user does | Where it runs | Permissions | Manifest | Admin |
|---|---|---|---|---|---|
| **Manual** (recommended v1) | Ticks emails in the pane and presses "Criar rascunhos (n)" | Pane plus backend AI; drafts created by the pane | Delegated `Mail.ReadBasic`, then `Mail.ReadWrite` | None | Consent for the delegated scopes |
| **Assisted** (v2) | The pane pre-ticks suggestions (rules from (iv-a), Focused only, optionally AI triage of subject and sender); the user confirms | Client-side version: pane plus backend. Server-side version: F-013 | Client-side: as manual. Server-side: F-013 app-only | None | Server-side: as F-013 |
| **Automatic** (F-013) | Nothing; drafts appear in Drafts | Server (ADR-0005) | App-only, RBAC-scoped, never `Mail.Send` | None | App-only consent, RBAC for Applications, DPO |
| Manual via category (variant of C) | Assigns a category in any Outlook client | Server (F-013) | As automatic | None | As automatic |

**Recommendation**
1. Ship **manual** first as the F-008 plus F-009 features, with checkboxes and the style profile. It delivers both requests while keeping every email action user-initiated. It needs no stored email, no app-only access and no webhooks, and it can go ahead **before** ADR-0005 is decided.
2. F-013 then becomes the **automatic** mode. It reuses the same draft-generation backend operation, the same style profile and the same rules (iv-a), so the server-side work only adds the trigger (notifications) and the app-only Graph client.
3. **Assisted** with AI triage comes last, because it sends subjects of many emails to the AI that the user did not pick.

**Safeguards that apply to all modes** (proposals for the specs):
- **Draft only.** `createReply` plus `PATCH`; never `/send`, `/reply` or `/replyAll` (the single-call reply sends), never `Mail.Send`.
- **No duplicate drafts.** Before creating a draft, check whether the add-in already made one for that message. The mechanism is unverified (a custom property or extension on the draft, or a lookup in Drafts by `conversationId`).
- **Caps.** At most N per batch (open question U2), sequential processing with per-row status, and cancel between items.
- **Server modes only: check ownership of IDs.** If the pane ever sends message IDs to the backend and the backend acts with app-only rights, the backend must only touch `/users/{oid}/messages/{id}`, with `oid` taken from the validated token. Otherwise one user could make the server act on another mailbox (an ID-substitution risk).
- **Email text is untrusted** in every mode. The AI never decides that an email is answered or ignored without the user (manual and assisted), and in automatic mode it only produces drafts.

---

## 8. Required permission, manifest and configuration changes

| Change | Manual (B) | Multi-select (A) | Category (C) / Automatic |
|---|---|---|---|
| `manifest.xml` `<Permissions>` | `ReadItem` (no change) | `ReadWriteMailbox` | No change |
| `manifest.xml` elements | None | `SupportsMultiSelect`, `SupportsPinning`, possibly `DefaultMinVersion` 1.13/1.15, `<Version>` bump | None |
| Admin re-upload of the manifest | No | **Yes** (and admin-only install) | No |
| Entra app registration: API permissions | Add delegated `Mail.ReadBasic` and `Mail.ReadWrite` (and `Mail.Read` if bodies are needed without drafts); admin consent per tenant policy | Delegated `Mail.ReadWrite` only if drafts | Application permissions per ADR-0005 |
| Exchange | None | None | RBAC for Applications scoped to the group (ADR-0005) |
| Supabase | Style profile and rules tables; the function needs database access (new secret: a service-role or scoped database credential) | Same | Plus the F-013 state |
| `security.md` | Must change (see section 9) | Must change | Must change |

**Note:** F-002/AC-3 says "No Graph scope appears anywhere in `src/`". A Graph feature amends that on purpose, and its spec must say so.

---

## 9. Privacy implications and proposed `security.md` updates

**Implications by component**
- **OpenAI:** every option sends email text the user picked. That is covered by D5, which is still open. The style-profile build sends up to about 20 × 1,500 characters once. Prefer `/v1/chat/completions`, which keeps no application state (abuse logs up to 30 days), or the Responses API with `store: false`. EU residency needs an OpenAI approval track (open question P4).
- **Supabase:** the backend stops being stateless.
  - The minimal design (style profile plus rules) stores **no email content**, only derived or user-confirmed data, keyed by `oid`, in eu-west-1.
  - Few-shot (ii) would store email content and is deferred.
  - Erasure is complete only after the backup window (plan-dependent, **unverified**).
- **Microsoft 365:**
  - Drafts created by the add-in live in the user's own mailbox and are subject to DIMSCALE's normal retention.
  - Delegated access means the add-in can only reach the signed-in user's mailbox (Graph docs: delegated access covers the signed-in user, or folders shared with them).
- **Third parties:** correspondents' text passes through OpenAI in every option. Only (ii) and (iv-b) keep it.

**Proposed edits for `security.md`** (not made; for the orchestrator):
1. The "Graph" bullet: list the delegated scopes per feature (`Mail.ReadBasic` for the list, `Mail.ReadWrite` for drafts) and keep "`Mail.Send` is never requested".
2. Replace "Stateless. Stores no email content" (overview, backend responsibilities) with a precise statement: what is stored (profile, rules), keys, region, retention, the deletion path, and that email bodies are **not** stored.
3. Add a "Learning data" section: no fine-tuning; no stored email text without an approved DPIA; profile output validation; user review before saving.
4. Add a "Batch actions" rule: caps, sequential processing, no `/send`, `/reply` or `/replyAll` calls, an ownership check on IDs in server modes.
5. Secrets: the database credential the function needs becomes a secret in the Supabase secret store.

---

## 10. Proposed changes to other specs (proposals only, nothing edited)

**F-012 (task pane design)**
- **Tabs (Q6).** With a list view, the pane has at least two views, so the tab bar appears. Proposal: "Analisar" stays the active-item view, and "Responder" gains a sub-view "Selecionar e-mails" (the checkbox list). Alternatively, the list is the content of Wireframe 4 ("no message selected").
- **New wireframe at 320 px.** Each row: checkbox, sender name, subject (2 lines, `dir="auto"`, wraps), date, a "Focused" chip when relevant. Below the list:
  - "Selecionados: n de N";
  - a primary "Criar rascunhos (n)" button, disabled when n = 0 or above the cap;
  - progress and status per row, with text for each state, never color alone;
  - the existing "Sem envio automático" line.
  - Use Fluent `Checkbox` or a selectable `Table` with `aria-checked` and full keyboard support (F-012 AC-19 to AC-21).
- **Wording.** "Criar rascunhos" becomes **truthful** with Graph `createReply` (F-012 Q9 and the slot table say "Criar rascunho" needs F-009). Keep "Abrir resposta para revisão" for the single-email F-005 path.
- **Out-of-scope conflict.** F-012 excludes "any persisted preference" and a settings menu. A style profile and ignore rules are persisted per user (server-side). They need their own view ("Perfil de estilo": view, edit, delete) in a separate feature spec, and F-012 should note that the exclusion covers client-side storage only.
- **"Atividade IA".** Batch results stay in memory for the session only, so no activity log is introduced.

**F-013 (automatic drafts, file not on disk)**
- Frame F-013 as the **automatic mode** of a shared "reply drafts" capability. The manual mode (pane plus delegated Graph) is delivered first and is listed in ADR-0005 as an option considered. It needs none of ADR-0005's app-only, RBAC or webhook infrastructure.
- Reuse the backend draft operation, the style profile and the rules (iv-a). Add the category trigger (option C) as an optional selection method, pending verification of the `Updated` notification on category changes and of category `$filter` support in subscriptions.
- Add the duplicate-draft guard and the ownership check on IDs.

**Backlog (status.md)**
- F-008 becomes "List Inbox emails with checkboxes (Graph `Mail.ReadBasic`)". `Mail.ReadBasic` replaces `Mail.Read` for the list.
- F-009 becomes "Create reply drafts for ticked emails (Graph `createReply` + `PATCH`, `Mail.ReadWrite`)".
- New IDs to assign (the orchestrator numbers them; F-014 is tentatively claimed by F-012 Q8): "Style profile from ticked sent replies", "Reply/ignore rules", and later "Few-shot examples" (only if approved).
- New decision: "Selection mode for v1 (manual vs multi-select vs category)" (U3).

---

## 11. Risks
- **Privilege creep.** Choosing A raises the manifest to the highest permission ("sending items"), even though the code never sends. Choosing B re-adds delegated `Mail.ReadWrite`, which the admin is about to revoke as unused. Both need a clear justification to the admin.
- **A stateful backend.** The first database use brings a new secret, new failure modes and new data subject rights. Keep the schema tiny.
- **Over-sharing in the profile.** The model might copy names or phrases with personal details into the profile. Mitigate with schema validation, rejection patterns and user review.
- **Batch cost and latency.** N emails means N model calls. A cap and sequential progress are required, and the user can cancel.
- **Expectation gap.** "Learn" suggests training. The product should say "perfil de estilo" and show the profile.
- **The EWS shutdown window** (October 2026 to April 2027) could affect Office.js APIs that use EWS IDs (option A only; undocumented).
- **Spec drift.** ADR-0005 and F-013 were not readable here; section 10 may conflict with their final content.

---

## 12. Items the documentation does not confirm (unverified)
1. The `SupportsMultiSelect` vs `SupportsMultiselect` casing, and whether `DefaultMinVersion` must be raised (A).
2. `loadItemByIdAsync` behavior after Exchange Online disables EWS (A).
3. `displayReplyFormAsync` on a `LoadedMessageRead` (listed, behavior undocumented) (A).
4. Text body format for `GET /me/messages/{id}` with `Prefer: outlook.body-content-type="text"` (B).
5. Which properties `Mail.ReadBasic` returns (`inferenceClassification`, `isRead`, `hasAttachments`): inferred from the exclusion list (B, iv-a).
6. That the add-in activates on Sent Items read forms (i, no-permission variant).
7. That changing a category triggers an `Updated` change notification (inferred), and that subscriptions accept `$filter` on `categories` (Q&A only) (C).
8. Supabase: whether deleted rows persist in backups until they expire, the backup region, and DIMSCALE's plan (i, ii).
9. The mechanism to detect an existing add-in draft for a message (duplicate guard).
10. OpenAI chat-model pricing for the batch and profile calls was not looked up (only the embeddings illustration). Fine-tuning pricing was not checked (moot).

---

## 13. Open questions

**For the user**
- **U1.** `requirements.md` excludes "bulk actions" and "automatic rules or filters". Do you approve amending it for (a) selecting several emails for drafts and (b) user-confirmed ignore rules?
- **U2.** Maximum emails per batch? Proposal: 10 (cost and time bound), later 20.
- **U3.** Selection method for v1: a list in the pane (recommended), Ctrl+click in Outlook's own list (needs the highest add-in permission and an admin re-upload, no body reading on Mac), or categories (needs F-013 first)?
- **U4.** List scope: Inbox only? Focused only by default? Last 7 days?
- **U5.** Is the list acceptable without the preview text? Previews need `Mail.Read` instead of `Mail.ReadBasic`.
- **U6.** Learning v1 = style profile from your own sent replies, profile shown and editable, no emails stored. Agree?
- **U7.** Would a "respostas modelo" library (texts you write yourself) cover the "learn content" wish instead of storing real emails (option ii)?
- **U8.** Is v1 order manual → (assisted) → automatic acceptable, with F-013 reusing the same draft engine?
- **U9.** D2 (which Outlook clients): does anyone use Mac or mobile? This matters only if option A or C is chosen.

**For the tenant admin (MMSI)**
- **A1.** The pending request is to revoke the old delegated `Mail.ReadWrite`. If the manual mode is approved, the app registration would request delegated `Mail.ReadBasic` and `Mail.ReadWrite` on purpose. Is that acceptable, and will you grant tenant-wide consent (users may not be allowed to consent themselves)?
- **A2.** Only if option A is chosen: is an add-in with `ReadWriteMailbox` acceptable? Are users on Current Channel at Version 2412 or later?
- **A3.** Only for option C or the automatic mode: as in ADR-0005 (app-only permissions, RBAC for Applications, group scope).

**For the DPO**
- **P1.** D5 still applies: sending the text of user-selected emails to OpenAI (DPA, lawful basis, notice line).
- **P2.** Storing a derived per-user style profile (no email text) in Supabase eu-west-1: acceptable? What retention, and what deletion on offboarding?
- **P3.** Storing real email text and embeddings for few-shot use (option ii): is a DPIA required? It is deferred unless approved.
- **P4.** Should DIMSCALE seek OpenAI EU data residency? It requires Zero Data Retention, Modified Abuse Monitoring or a similar approval.
- **P5.** Is it acceptable that a deleted profile or rule persists in database backups for the plan's backup window (for example 7 days on Pro, if that is the plan)?
- **P6.** The offboarding process: who triggers deletion of a leaver's profile and rules?

---

## 14. Sources consulted
**Microsoft Learn**
- https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select
- https://learn.microsoft.com/javascript/api/outlook/office.mailbox?view=outlook-js-1.15 (`getSelectedItemsAsync`, `loadItemByIdAsync`, `addHandlerAsync`/`SelectedItemsChanged`, `convertToRestId`, `convertToEwsId`, `displayMessageForm`, `displayMessageFormAsync`, `makeEwsRequestAsync`)
- https://learn.microsoft.com/javascript/api/outlook/office.loadedmessageread
- https://learn.microsoft.com/javascript/api/manifest/action (`SupportsMultiSelect`, `SupportsNoItemContext`, `SupportsPinning`)
- https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets
- https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions
- https://learn.microsoft.com/office/dev/add-ins/concepts/privacy-and-security
- https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0
- https://learn.microsoft.com/graph/permissions-reference (`Mail.ReadBasic`, `Mail.Read`, `Mail.ReadWrite`)
- https://learn.microsoft.com/graph/api/resources/message?view=graph-rest-1.0
- https://learn.microsoft.com/graph/api/resources/mailfolder?view=graph-rest-1.0
- https://learn.microsoft.com/graph/api/message-createreply?view=graph-rest-1.0
- https://learn.microsoft.com/graph/outlook-change-notifications-overview
- https://learn.microsoft.com/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online
- Microsoft Q&A threads on category filters (not authoritative; used only to mark item 12.7 as unverified): https://learn.microsoft.com/answers/questions/1418833/microsoft-graph-filter-messages-by-categories

**OpenAI** (platform.openai.com redirects to developers.openai.com)
- https://developers.openai.com/api/docs/guides/your-data
- https://developers.openai.com/api/docs/guides/supervised-fine-tuning
- https://developers.openai.com/api/docs/guides/model-optimization
- https://developers.openai.com/api/docs/guides/embeddings

**Supabase**
- https://supabase.com/docs/guides/database/extensions/pgvector
- https://supabase.com/docs/guides/database/vault
- https://supabase.com/docs/guides/database/extensions/pgsodium
- https://supabase.com/docs/guides/platform/backups
- https://supabase.com/docs/guides/security/platform-security (it does not state the at-rest algorithm or region; the pgsodium page was used for "encrypted at rest by default")

GDPR principles mentioned (minimisation, storage limitation, erasure, processors, DPIA) are general context for the DPO, not legal advice, and were not checked against a legal source in this session.
