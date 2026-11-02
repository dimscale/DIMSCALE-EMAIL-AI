# Research: F-010 compose-mode support (help writing and improving an email being composed)

- **Date:** 2026-11-02. **Author:** requirements-architect. **Status:** research only, no spec yet. Nothing here is approved; the user decides the items in "Decisions for the user".
- **Question:** what Office.js and manifest changes does it take to read and write the body of an email the user is composing, what does that force on the tenant admin and the users, which Outlook clients support it, and should it ride in the manifest version now going to MMSI?

## Short answer
1. **Reading** the compose body works with the current permission (`ReadItem`): `body.getAsync` is Mailbox 1.3, read item, Compose or Read.
2. **Writing** into the compose body (`setAsync`, `prependAsync`, `setSelectedDataAsync`) needs **`ReadWriteItem`**, which in the add-in-only manifest is a **top-level `<Permissions>` change that applies to the whole add-in**, including the read pane.
3. **Either way the manifest changes**: the pane cannot appear in compose without a new `MessageComposeCommandSurface` extension point. **A manifest change means a new `<Version>` and the tenant admin must upload the new manifest** (line-of-business add-in: "the admin must upload the new manifest file" and select **Update**). Users get it the next time they start Outlook. Microsoft does not document whether a permission increase asks the admin for anything extra (open question O-2).
4. **Recommendation: defer F-010 to a second manifest update**, and do not put an unbuilt compose surface or `ReadWriteItem` in the manifest that goes to MMSI now (reasons and cost in "Recommendation").

## 1. Office.js: what reads and writes the compose body
All on the `Office.Body` interface (`item.body`) of `MessageCompose`. Source: https://learn.microsoft.com/javascript/api/outlook/office.body (checked 2026-11-02, preview moniker page; member requirement sets are the stable ones listed).

| Member | Requirement set | Minimum permission | Mode | Notes from the page |
|---|---|---|---|---|
| `getAsync(coercionType, [options], callback)` | Mailbox **1.3** | **read item** | Compose or Read | "Returns the entire current body in the format specified by `coercionType`." In Outlook on the web and new Outlook on Windows `Text` "may not always provide reliable formatting"; use `Html` for formatted bodies. `bodyMode` applies only to replies on the Compose surface and is ignored in classic Windows and Mac; in web and new Outlook the default `FullBody` returns the whole conversation thread. |
| `getTypeAsync` | Mailbox 1.1 | read item | Compose | Recommended before writing: pass the returned type as `coercionType` (the body is plain text or HTML). |
| `setAsync(data, options, callback)` | Mailbox **1.3** | **read/write item** | Compose | Replaces the whole body; 1,000,000-character limit (`DataExceedsMaximumSize`); `InvalidFormatError` if `coercionType` is `Html` and the body is plain text. **In classic Windows and Mac the user cannot undo it with Undo.** In web and new Outlook, with `bodyMode` missing or `FullBody`, the whole body including earlier conversation messages is replaced (a notification tells the user and the grouping setting is restored after the call); `HostConfig` honours the user's grouping. Cursor ends at the start (classic Windows), the end (web, new Outlook) or is lost (Mac). Not supported on a message loaded with `loadItemByIdAsync`. Not supported in Message Compose on Android and iOS. |
| `prependAsync(data, options, callback)` | Mailbox 1.1 | read/write item | Compose | Inserts at the start of the body; 1,000,000-character limit; cursor position behaviour differs by client (kept in web and Windows, lost on Mac); not supported in Message Compose on Android and iOS. |
| `setSelectedDataAsync(data, options, callback)` | Mailbox 1.1 | read/write item | Compose | "Replaces the selection in the body": inserts at the cursor or replaces selected text; if the cursor was never in the body or the body lost focus (clicking in the pane does that), "the string will be inserted at the top of the body content". Cursor ends after the inserted text. |
| `appendOnSendAsync`, `prependOnSendAsync` | 1.9 / 1.13 | read/write item + extended permission | Compose | **Not proposed**: they act at send time, which is outside "the user controls every email action". |

Not verified in this research (to check on Microsoft Learn before a spec): reading the selected text (`item.getSelectedDataAsync`), `item.subject.getAsync/setAsync` in compose, how `Html` coercion treats the signature and the quoted reply, and Mailbox 1.3 wording for `setAsync` coercion on each client.

## 2. Manifest changes
Current manifest (`manifest.xml`, `<Version>` 1.0.2.0): top-level `<Permissions>ReadItem</Permissions>`, a top-level `<Rule>` `ItemIs Message FormType="Read"`, a `FormSettings` `ItemRead` form, and one `MessageReadCommandSurface` button (`ShowTaskpane`, `SupportsPinning`).

| Change | What it takes | Source |
|---|---|---|
| Show the pane or a button in compose | Add `<ExtensionPoint xsi:type="MessageComposeCommandSurface">` in `DesktopFormFactor` (child `OfficeTab` or `CustomTab`; "puts buttons on the ribbon for add-ins using mail compose form"), with its own `Control`, `Resources` strings and icons. | https://learn.microsoft.com/javascript/api/manifest/extensionpoint |
| Form type | `Form` `xsi:type="ItemEdit"` is for "when a user composes a reply, creates a new message or appointment". The existing file marks `FormSettings` as "ignored when VersionOverrides is present" (project note, not a Microsoft quote). **Whether the top-level `<Rule>` (`FormType="Read"`) and `FormSettings` must also gain an `Edit`/`ReadOrEdit` entry for compose to activate is not clear from the pages read** (the Rule page: `FormType` values `Read`, `Edit`, `ReadOrEdit`; within an `ExtensionPoint` it must be `Read`). Open question O-1; settle with `office-addin-manifest validate` and a sideload test. | https://learn.microsoft.com/javascript/api/manifest/form, https://learn.microsoft.com/javascript/api/manifest/rule |
| Permission for writing | `<Permissions>ReadWriteItem</Permissions>` (top level, **whole add-in**). `ReadItem` cannot call `body.setAsync`, `prependAsync` or `setSelectedDataAsync`. `ReadWriteItem` adds "full Outlook add-in API access except `makeEwsRequestAsync`" and "setting the item properties" (and attachments add/remove). It does **not** allow sending (that is `ReadWriteMailbox`, never requested). | https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions |
| Read-only compose option | With `ReadItem` the pane can read the compose body (`getAsync`) and show an improved text for the user to copy; no write. The compose extension point is still needed. | same page, Body page above |
| Pinning in compose | Pinnable task panes are supported in Message Compose and Read, **but pinning is not shared across modes** (the user pins separately in compose). | https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane |
| Version and resources | `<Version>` must be bumped (1.0.2.0 to a higher value) so the update is recognised. | https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins |

Least privilege (CLAUDE.md rule 7): `ReadWriteItem` is the lowest level that can write the body, but it is a real escalation for the read feature set too. If only "show an improved text, user copies it" is built, `ReadItem` is enough.

## 3. Admin upload and user impact
- **Does a manifest change force a new upload? Yes.** Microsoft: for a line-of-business add-in "where an admin uploaded a manifest, the admin must upload the new manifest file to support metadata changes. The next time a user starts the relevant Microsoft 365 productivity application, the add-in updates." The admin selects the add-in and chooses the **Update** button. (Microsoft Marketplace add-ins update without admin action; ours is not one.) Source: https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins. Code and UI changes (the page at `https://dimscale.github.io/DIMSCALE-EMAIL-AI/`) do **not** need the admin; only the manifest does.
- **Does a permission increase need extra consent?** The same page says that when a manifest is updated "typical changes are to an add-in's icon and text... the permissions of the add-in don't change", so a permission change is outside what it describes. A web search found no Microsoft page for a `ReadItem` to `ReadWriteItem` update; third-party reports suggest re-approval, which is **not authoritative**. Treat it as unknown (O-2) and ask MMSI.
- **User impact:** users restart Outlook to get the update; a new compose button appears in compose windows only; **Microsoft says new add-ins can take 24 to 72 hours to appear on the ribbon** (stated for new deployments, not specifically updates). Existing users see no change in the read pane. An in-pane notice that compose text goes to OpenAI is needed (below).
- **Sources for the MMSI side already in the repo:** `docs/admin/2026-10-23-mmsi-request.md` and F-012 AC-52/AC-53 (the final manifest goes to the admin only after the live ribbon check).

## 4. Platform support
- **Requirement sets.** The members above are Mailbox 1.1 and 1.3, below the 1.9 that F-003 already needs on read. From the F-003 research of the requirement-set page (https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets): the new Mac UI supports 1.1 to 1.14; the **classic Mac UI supports only 1.1 to 1.8**; classic Windows (Microsoft 365) supports 1.9 from Version 2008; new Outlook on Windows and modern Outlook on the web support the later sets. 1.1 and 1.3 are therefore expected everywhere the add-in runs; confirm in the spec with `isSetSupported('Mailbox','1.3')`.
- **Add-in commands in compose** are available in Outlook on the web, new Outlook on Windows, Outlook 2016 or later on Windows and Outlook on Mac. Mobile does not support `prependAsync` or `setAsync` in Message Compose (not a DIMSCALE target). Source: https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands and the Body page.
- **Per client differences that affect UX** (Body page): no Undo after `setAsync` in classic Windows and Mac; whole-thread replacement in web and new Outlook unless `bodyMode: HostConfig`; cursor behaviour differs; `getAsync(Text)` formatting less reliable in web and new Outlook (use `Html`); `setSelectedDataAsync` inserts at the top if the body lost focus.
- **Activation limits** apply as for read: no activation on IRM-protected or encrypted items in some clients, shared mailboxes need the shared-folder permission, etc. (https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview).
- **Live checks needed (nothing here is tested):** all three clients (classic Windows, new Outlook, Mac), each with a new message, a reply and a forward; signature and quoted thread behaviour; plain-text versus HTML compose; undo behaviour; pinning in compose.

## 5. Security implications against the non-negotiable rules
| Rule | F-010 position |
|---|---|
| 1. The user controls every action; never send | Writing text into the compose body is allowed (the user still presses Send), as for the reply form. The code never calls send or `appendOnSendAsync`/`prependOnSendAsync`. `ReadWriteItem` does not include sending, and `Mail.Send` or `ReadWriteMailbox` are never requested. Write only after an explicit click on a clear button ("Inserir no e-mail"), never automatically, and never on text the user did not see. Because Undo is missing in classic Windows and Mac after `setAsync`, show the previous text in the pane ("Versão anterior") before replacing, or prefer `setSelectedDataAsync`/`prependAsync` over replacing the whole body. |
| 2. No secrets on the client | Unchanged. OpenAI called only from `analyze-message`. |
| 3. Email is untrusted | The text being composed is the **user's own** text, but a reply compose body can contain the quoted original (untrusted) and the AI must not follow instructions in either. Same delimiters, fixed instructions, no tools, strict schema. AI output is rendered in the pane as text nodes and inserted into the body as **plain text** (`coercionType: Text`, or HTML-escaped when the body is HTML), never as markup from the model; the existing "no `<`, `>`, link or address" output checks should apply (a draft the user asked to improve may legitimately contain links and addresses, which is a spec decision, O-3). |
| 4. Minimal data | Send only the user's draft text (up to a limit to be set in the spec), and, for improvement, the subject. Exclude recipients (To/Cc/Bcc), signatures and, by default, the quoted history. Reading with `bodyMode: HostConfig` or stripping the quoted part is needed so the whole thread is not sent unintentionally (web and new Outlook default to `FullBody`). A user's unsent draft can contain more sensitive content than an incoming mail; this is a **new data flow for D5 (DPO)**: unsent drafts and possible third-party text go to OpenAI on a click. |
| 5. No logging of content | Unchanged rule; the draft text is never logged by the pane or the function. |
| 6. Don't invent APIs | Members above are confirmed on Microsoft Learn; items listed as "not verified" must be checked before the spec. |
| 7. Least privilege | `ReadItem` for a read-only "improve and copy" variant; `ReadWriteItem` only if the body is written, and it then applies to the whole add-in. |
| 8. No new platform | No new platform needed; same function (a new action such as `improve`, or an extension of `reply`'s refinement) and the same caps. |

**New data-flow facts to put in security.md when specified:** a new trigger (the compose pane), a new input (the user's draft, possibly with quoted third-party text), and for the write path a manifest permission change. Prompt injection from the quoted text could produce misleading suggested text; the user reviews before inserting and sends.

## 6. Recommendation: defer to a second manifest update
**Recommendation (architect, for the user to decide): do not put F-010 in the manifest version now going to MMSI. Send the current manifest (1.0.2.0, `ReadItem`, read surface) and plan F-010 as a second manifest update.**

Reasons:
1. **Nothing to ship in it.** F-010 has no spec, no plan and no code. A `MessageComposeCommandSurface` button without a working pane is a dead control, which F-012 forbids ("no disabled fake control"), and `ReadWriteItem` without code that uses it breaks least privilege (rule 7).
2. **It would widen what MMSI is asked to approve** for the whole add-in (`ReadWriteItem` is not part of the MMSI request now, which asks for `ReadItem` and `Mail.ReadBasic`) and could add review questions or re-consent that delay the first, valuable upload. The effect of a permission increase on approval is undocumented (O-2).
3. **Untested surface.** Compose activation (Rule and FormSettings, O-1), pinning in compose and per-client body behaviour are all unverified; F-012 AC-52 already demands a live ribbon check before the admin upload, and compose would add three clients x several compose types to it.
4. **New DPO scope.** Sending the user's unsent drafts to OpenAI is a different processing from F-003 and should be raised with the DPO (D5) before it is wired in.

**Cost of the second upload (what deferring costs):**
- One more request to the tenant admin (MMSI): upload the new manifest and press **Update** (documented). Turnaround time with MMSI is **unknown** (the first request will show it).
- Users restart Outlook once; the compose button may take up to 24 to 72 hours to appear (documented for new deployments).
- Possibly an extra approval step for the `ReadWriteItem` increase (undocumented, O-2).
- Version bump (for example 1.1.0.0) and a repeat of the live ribbon check for the new surface.
- No change needed for the read pane; code deploys through GitHub Pages without the admin.

**When including it now would make sense:** if the user needs F-010 within a few weeks, **and** MMSI's turnaround is slow, **and** the spec is approved and the compose pane is built and sideload-tested in the three clients **before** the MMSI request is sent (the repo's rule is that the final manifest is sent once, after the live check). The cheaper compromise inside that case is a **read-only compose variant** (`MessageComposeCommandSurface` plus `ReadItem`, text shown for copy), which needs no permission escalation and could be extended to writing later through the same second upload.

## Decisions for the user
1. **Manifest timing:** defer F-010 to a second manifest update (recommended), or include a compose surface in the current manifest only after F-010 is specified, built and sideload-tested in three clients?
2. **Write or read-only:** is writing into the compose body required (needs `ReadWriteItem` for the whole add-in), or is "improve, then copy" enough (`ReadItem`)?
3. **What may be sent:** only the draft body? the subject too? the quoted history never or optionally? any limit (the spec would propose 4,000 characters like F-003)?
4. **Links and addresses in improved text:** the F-003 output rule rejects any link or address in AI output; a draft the user wants improved may contain them. Keep the strict rule (then drafts with links fail or have them removed) or define a rule that only allows links and addresses already present in the user's own text?
5. **DPO:** add the user's unsent drafts (and quoted third-party text) to the D5 question to the DPO now.
6. **Ask MMSI:** does raising the permission from `ReadItem` to `ReadWriteItem` in an updated manifest require anything beyond the admin's Update (O-2)? (Cheap to ask in the first request.)

## Open questions
- **O-1.** For a mail add-in with a `VersionOverrides` `MessageComposeCommandSurface`, must the top-level `<Rule>` (currently `FormType="Read"`) and `FormSettings` also declare an edit form for compose activation? The pages read do not say. Settle with the manifest validator and a sideload test.
- **O-2.** Does a `ReadItem` to `ReadWriteItem` increase in an updated centralized-deployment manifest need extra admin approval or user consent? Not documented in the pages read.
- **O-3.** Which Output rules apply to improved text (see decision 4)?
- **O-4.** How should the pane read only the user's new text and not the quoted thread in each client (`bodyMode: HostConfig` is ignored in classic Windows and Mac; whether `Text` contains the quoted history there is not stated)? Live check.
- **O-5.** Does `body.getAsync(Html)` plus `setAsync(Html)` preserve the user's signature and formatting? Live check.
- **O-6.** Is the selected text readable (`item.getSelectedDataAsync`) for an "improve the selection" mode? Not verified here.
- **O-7.** Is an unsent draft included in the OpenAI retention terms the DPO reviews (abuse-monitoring logs up to 30 days, `store: false`)? Same OpenAI terms as in security.md; DPO to confirm for drafts.

## Sources consulted (2026-11-02)
- https://learn.microsoft.com/javascript/api/outlook/office.body (Body members, requirement sets, permissions, per-client notes)
- https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions (restricted, read item, read/write item, read/write mailbox)
- https://learn.microsoft.com/javascript/api/manifest/extensionpoint (MessageComposeCommandSurface)
- https://learn.microsoft.com/javascript/api/manifest/form (ItemRead and ItemEdit)
- https://learn.microsoft.com/javascript/api/manifest/rule (ItemIs `FormType`)
- https://learn.microsoft.com/office/dev/add-ins/develop/create-addin-commands (add-in commands, Outlook client list)
- https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane (pinning in compose and read, client versions)
- https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview (supported clients, activation limits)
- https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins (admin must upload a new manifest for a line-of-business add-in; 24 to 72 hours for the ribbon)
- https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets (client support by set, as recorded in F-003 on 2026-10-29, not re-fetched today)
- Repo: `manifest.xml`, `docs/specs/F-003-analyze-summary-and-reply.md`, `docs/architecture/security.md`, `docs/status.md`, `docs/product/requirements.md`
