# F-012: Task pane design (Outlook-native, 320 px compact, "DIMSCALE MAIL" ribbon entry)

- **Status:** Approved (2026-10-23) with the user's choices: **direction D "Outlook nativo"** (Fluent stock themes plus the user's Office theme colours; only the wordmark and the name are DIMSCALE), **320 px compact** (only the information the user needs), and a **"DIMSCALE MAIL" ribbon entry** that opens a pinnable pane. The "Responder" tab hosts the contact selection for automatic drafts (F-013). Earlier directions A, B and C are dropped.
- **Amended 2026-10-23** (requirements-architect) to match that approval. Superseded and removed: directions A/B/C, the DIMSCALE navy/teal Fluent ramps, the brand contrast table, the full-layout wireframe with future AI slots, the future-slot components, and "no manifest change". AC IDs are kept; superseded ACs are marked.
- **Amended again 2026-10-23** (requirements-architect) with the user's answers and one orchestrator decision: button label "DIMSCALE MAIL" (Q-R2 closed); the user wants an own "DIMSCALE MAIL" tab (Q-R1 becomes the experiment ACs 37 to 40, production manifest unchanged until it passes); clients in use are classic Windows, new Outlook and Mac (D2: Mac facts and live checks added); pinned-pane behaviour clarified (AC-36, AC-41); the message excerpt is **collapsed by default** in a "Mensagem" accordion (AC-10, AC-12).
- **Amended 2026-10-29 for the Kutools-style pane (user decisions of 2026-10-29, relayed by the orchestrator; requirements-architect).** The user's decisions: four tabs **Resumo | Resposta | Traduzir | Automático** (Automático = the former Responder tab), each tool a card with copy and expand actions; ribbon group "DIMSCALE MAIL" with an **"IA"** button and a **"Resumo"** button, changed **before** anything goes to the tenant admin; only Portuguese and English; no model selector, other languages, "Email Learning" (F-016) or reading-pane toggles. The AI behaviour (auto mode, translate, reply options) is specified in [F-003](F-003-analyze-summary-and-reply.md). **Superseded by this amendment:** the Analisar/Responder tab pair (AC-16), the production content set wireframe (AC-10), the single "DIMSCALE MAIL" button (AC-32), "no new Office.js call except theme/`ItemChanged`" (AC-2). Items marked **(architect proposal)** below are choices made while writing this amendment that the user has not seen yet; they are listed in "Open questions" (Q-K1 to Q-K6) and need the user's confirmation. The approval of 2026-10-23 stays on record; the user confirms the amended text.
- **Tier:** Active item, presentation, plus the manifest (ribbon entry, pinning). No Graph, no new add-in permission (`ReadItem` stays). Since 2026-10-29 the pane also hosts the F-003 AI cards and stores one per-user preference with `Office.context.roamingSettings` (in the user's own mailbox, F-003).
- **Depends on:** F-001, F-002, F-011 (all Done).
- **Plan:** [docs/plans/F-012-taskpane-design.md](../plans/F-012-taskpane-design.md)

## Goal
A compact pane that looks like part of Outlook, opened from a "DIMSCALE MAIL" ribbon entry, pinnable, that:
1. keeps everything F-001 and F-002 do (message context, backend status);
2. follows the user's Outlook theme (light/dark) at 320 px;
3. ~~has two tabs: **Analisar** and **Responder**~~ (superseded 2026-10-29) has four tabs: **Resumo**, **Resposta**, **Traduzir** (AI cards for the open message, F-003) and **Automático** (contact selection for automatic drafts, F-013, formerly "Responder").

## Design principles (kept)
1. Content over chrome. One compact header row; no brand fills, no brand accent colour.
2. The user is in control and the UI says so: "Sem envio automático" is always visible on Analisar. No label promises something not built.
3. Email content is untrusted: rendered only as React text nodes.
4. Show only what works: an unbuilt feature is one short "em breve" card, never disabled fake controls.
5. Accessible by construction: WCAG 2.1 AA, keyboard, screen reader names, no meaning by colour alone, reflow at 320 px.
6. Native components: Fluent UI React v9, stock themes.

## Verified platform facts
| Fact | Source |
|---|---|
| Task pane best practices; Outlook on the web pane size 320x570; favor content over chrome. | https://learn.microsoft.com/office/dev/add-ins/design/task-pane-add-ins, https://learn.microsoft.com/office/dev/add-ins/design/add-in-design |
| Accessibility: keyboard, ARIA, no colour-only meaning, 4.5:1 text / 3:1 non-text, four Windows contrast themes, Narrator/JAWS/NVDA. | https://learn.microsoft.com/office/dev/add-ins/design/accessibility-guidelines |
| `Office.context.officeTheme` in Outlook: Mailbox **1.14**, task panes only (not event-based). `bodyBackgroundColor`, `bodyForegroundColor`, `controlBackgroundColor`, `controlForegroundColor` are hex triplets. `isDarkTheme` and `themeId` are **not supported in Outlook** ("use the `body*Color` and `control*Color` properties instead"). Theme change: `Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, handler)`, Mailbox 1.14. | https://learn.microsoft.com/javascript/api/office/office.context#office-office-context-officetheme-member, https://learn.microsoft.com/javascript/api/office/office.officetheme, https://learn.microsoft.com/javascript/api/outlook/office.mailbox, https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-14 |
| Classic Outlook on Windows has Mailbox 1.14 from Version 2404; NAA (F-002) needs 2409, so every client that can sign in has 1.14. Check with `Office.context.requirements.isSetSupported("Mailbox", "1.14")`. | https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets |
| Fluent UI React v9 stock themes `webLightTheme`, `webDarkTheme` and the high-contrast theme exist in the installed `@fluentui/react-components` 9.74.10 / `@fluentui/tokens` (checked in `node_modules`; the Fluent docs site did not render to the fetch tool). | Installed package, https://learn.microsoft.com/office/dev/add-ins/quickstarts/fluent-react-quickstart |
| **CustomTab in Outlook:** "In Outlook on the web, on Mac, and in the new Outlook on Windows, the `<CustomTab>` element isn't available, but you can put *custom* groups of controls on one of the built-in OfficeTabs instead." Valid in Mail 1.0/1.1 schemas. Add-ins are limited to one custom tab. | https://learn.microsoft.com/javascript/api/manifest/customtab |
| "For Outlook, custom tabs are only supported in classic Outlook on Windows." When a command goes on a built-in tab, "the command will appear on the default tab based on the platform and current Outlook mode." | https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands |
| **Not documented:** what new Outlook, the web or Mac do with a manifest that declares a `<CustomTab>` (ignored, fallback, or install error). The "use `CustomTab` or `OfficeTab`, not both" rule is stated only for `PrimaryCommandSurface` (Word/Excel/PowerPoint/OneNote); for `MessageReadCommandSurface` both are listed as children with no statement on combining them. | https://learn.microsoft.com/javascript/api/manifest/extensionpoint |
| `OfficeTab` in Outlook: "only one OfficeTab element with the `TabDefault` ID is supported"; "You can add only one group per add-in to the default tab." Mail 1.1 schema → Mailbox 1.5. | https://learn.microsoft.com/javascript/api/manifest/officetab |
| `Group`: `Label` required (resid ≤ 32 chars); the group `Icon` is "not supported in Outlook add-ins". | https://learn.microsoft.com/javascript/api/manifest/group |
| "In Office on the web, if you're using the single-line or simplified ribbon layout, the add-in name isn't shown on the ribbon. Only the add-in's command icon is shown." | https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands |
| **Pinning:** `<SupportsPinning>true</SupportsPinning>` on the `ShowTaskpane` `<Action>`; it is in the VersionOverrides **v1.1** schema, so both v1.0 and v1.1 are needed. Available in modern Outlook on the web, new Outlook on Windows, classic Outlook 2016+ on Windows, Mac 16.13+; **not** for appointments or Outlook.com. Handler: `Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler)` in `Office.onReady`; when it fires "`Office.context.mailbox.item` is already updated"; the handler "should check whether or not the `Office.context.mailbox.item` is null". | https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane |
| Ribbon icons: PNG 16/32/80 for `bt:Image`; manifest `IconUrl` 64 and `HighResolutionIconUrl` 128 for mail add-ins. | https://learn.microsoft.com/office/dev/add-ins/design/add-in-icons, https://learn.microsoft.com/javascript/api/manifest/iconurl |
| **Mac (D2, checked 2026-10-23):** new Mac UI (Outlook 16.38.506 or later) supports Mailbox 1.1 to **1.14**, NestedAppAuth 1.1 and DialogAPI; the **classic** Mac UI supports only Mailbox 1.1 to **1.8** (so no `officeTheme` or `OfficeThemeChanged`, which are 1.14), although the table also lists NestedAppAuth 1.1 for it. Classic Outlook on Windows: 1.14 from Version 2404, 1.16 from 2602. New Outlook on Windows and modern web: 1.1 to 1.16 and NestedAppAuth 1.1. | https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets |
| Add-ins using the **unified manifest** can't be installed in Outlook on Mac, so this add-in stays on the add-in only (XML) manifest while Mac is a target. | https://learn.microsoft.com/office/dev/add-ins/outlook/compare-outlook-add-in-support-in-outlook-for-mac |
| **Pinning on Mac:** supported from Outlook on Mac Version 16.13 (18050300) (Mailbox 1.5 feature). By default a task pane closes when another message is selected; pinned panes stay. Pinning isn't supported across modes (a pane pinned in read mode isn't shown in compose) and isn't available for appointments or Outlook.com. Docs say nothing on whether a pinned pane reopens after Outlook restarts. | https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane |
| `OfficeThemeChanged`: per-platform firing is **not documented** beyond the requirement set (1.14: new Mac UI yes, classic Mac no). Live check on each client. | https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-14 |
| Custom tab position elements `InsertBefore`/`InsertAfter` are PowerPoint-only; `OfficeGroup` is not available in Outlook; add-ins get one custom tab; the tab `Label` resid is at most 32 characters. Custom tab placement markup "is only available in PowerPoint on the web". | https://learn.microsoft.com/javascript/api/manifest/customtab, https://learn.microsoft.com/office/dev/add-ins/design/custom-tab-placement |
| A manifest change needs a `<Version>` bump and an admin re-upload (F-011 runbook); code-only changes reach users without it. | [F-011](F-011-github-pages-hosting.md) |
| **Two task pane buttons (checked 2026-10-29).** `ShowTaskpane` children in Outlook: `SourceLocation` (required, resid ≤ 32 chars pointing to a `bt:Url`), `SupportsPinning` ("Supported in Outlook only", Mail 1.1 schema, Mailbox 1.5), `SupportsMultiselect`, `SupportsNoItemContext`. **`TaskpaneId` is "Not supported in Outlook add-ins"** and **`Title` is "Not supported in Outlook add-ins"**. The documented "same `TaskpaneId` → the pane container remains open but the contents … replaced with the corresponding Action `SourceLocation`" applies to Task pane 1.0 (Word/Excel/PowerPoint), **not** to Outlook. So Microsoft does **not document** what Outlook does when two `ShowTaskpane` buttons with different `SourceLocation`s are used (second pane, replace and reload, or no change). | https://learn.microsoft.com/javascript/api/manifest/action |
| **Shared runtime:** "Shared runtimes aren't supported in Outlook" and "You can share runtimes only in Excel, PowerPoint, and Word". So a function command cannot call `Office.addin.showAsTaskpane` to switch the pane to a tab in Outlook. A non-shared task pane runtime "shuts down when the task pane is closed". | https://learn.microsoft.com/office/dev/add-ins/testing/runtimes |
| `SourceLocation` URLs: `bt:Url` values must be HTTPS, at most 2048 characters; resource `id` at most 32 characters; `ShortStrings` at most 125 characters (labels, titles), `LongStrings` at most 250 (descriptions). A Microsoft example uses a query string in a task pane `bt:Url` (`https://localhost:3000?commands=1`), but **no page documents that Outlook passes a query string or hash through to the page**; a Microsoft Q&A thread (Word, unanswered, not authoritative) reports that `SourceLocation` query parameters were stripped from `location.search`. The URL **path** is the documented address of the page, so a second HTML page is the only entry signal that does not depend on undocumented behaviour. | https://learn.microsoft.com/javascript/api/manifest/resources, https://learn.microsoft.com/javascript/api/manifest/action, https://learn.microsoft.com/answers/questions/1669061/how-can-i-use-the-query-string-parameters-on-sourc (community) |
| Button: `Label`, `Supertip` (Title, Description; "Supertips are only supported in Office desktop clients"), `Icon` and exactly one `Action` are required; each control `id` unique in the manifest. With two or more commands in a group and limited space, Office shows small icons with labels, then icons only, then a dropdown; the client decides icon size (16 or 32 px), so a "big" button cannot be forced. Recommended at most seven commands per group; "Don't add unnecessary buttons". | https://learn.microsoft.com/javascript/api/manifest/control-button, https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands |
| Ribbon icons: Monoline style, at most two elements, flat colour with at most one extra colour, **no text characters** ("keep icons as language neutral as possible"), **PNG with a transparent background and 32-bit depth**; 16, 20 and 32 px are the 100 % sizes, 80 px for high DPI. | https://learn.microsoft.com/office/dev/add-ins/design/add-in-icons |
| **`Office.context.roamingSettings`** (used by F-003 auto mode): Mailbox **1.1**, minimum permission **restricted** (so `ReadItem` covers it), Compose or Read; "saved per add-in and per user … only from the user's mailbox in which they are saved"; values String/Number/Boolean/null/Object/Array; 32 KB per add-in; `set` then `saveAsync(callback)` persists. "Initialized from the persisted storage only when the add-in is first loaded … If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values … not available until the task pane … is closed and reopened." Classic Windows sees another client's change after the add-in reloads; web, Mac and new Outlook need a browser refresh or client restart. "Shouldn't be considered secure storage … shouldn't be used to store sensitive information". | https://learn.microsoft.com/javascript/api/outlook/office.roamingsettings, https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in |
| Fluent UI v9 `Switch`, `Dialog`, `TabList`, `Select` (a styled native `<select>`) and `Dropdown` are all exported by the installed `@fluentui/react-components` (`react-combobox` is nested under it). Option lists use `Select` (native, keyboard and screen-reader behaviour from the platform, no popup layer in a 320 px pane). | Installed packages (checked 2026-10-29 in `node_modules`) |

## Approved design (direction D "Outlook nativo")

### Theme
1. If `isSetSupported("Mailbox","1.14")` and `officeTheme.bodyBackgroundColor` is a valid hex triplet: dark when its relative luminance is below 0.179, else light. Use `webDarkTheme` or `webLightTheme`, and paint the pane body with `bodyBackgroundColor` and text with `bodyForegroundColor`, **only if** that pair has a contrast of at least 4.5:1; otherwise keep the Fluent theme's own neutral background and foreground (guard against odd Office colours).
2. Otherwise `matchMedia("(prefers-color-scheme: dark)")`, then light.
3. Re-evaluate on `OfficeThemeChanged` (registered when 1.14 is supported, removed on unmount) and on the `matchMedia` change event.
4. `forced-colors: active` uses Fluent's high-contrast theme and drops Office colours.
5. `isDarkTheme` and `themeId` are never used.

Brand: only the DIMSCALE wordmark (black on light, white on dark, per the brandbook) and the add-in icon (green "D" on white, user decision 2026-10-22). Buttons, links and focus use Fluent's stock brand tokens. No webfonts: Segoe UI (Fluent default) for everything; the wordmark is an inline SVG.

### Size and density (320 px compact)
Designed for 320 px, one column, vertical scroll only; works up to 500 px (column capped at 640 px). Padding 12 px, gaps 8 px, body 14 px, meta 12 px (floor), targets ≥ 32 px high. Email-derived text wraps (`overflow-wrap: anywhere`) and has `dir="auto"`.

### Content set (2026-10-29): four tabs, tool cards

#### Layout at 320 px
The header, footer, theme, density and backend status rules above and below are unchanged. Between header and footer:
- **Tab row:** Fluent `TabList` `size="small"`, labels **Resumo | Resposta | Traduzir | Automático**, text only, one row. Resumo is selected on open. It must fit at 320 px without horizontal scroll or truncation (AC-44). If the visual check shows it doesn't fit, the coder reports it and the user chooses a shorter label (Q-K4); the coder does not shorten labels on their own.
- **Message context** (Resumo, Resposta, Traduzir only; hidden on Automático): the F-012 subject, sender and date rows and the "Mensagem" accordion (collapsed by default), exactly as approved on 2026-10-23, rendered **once above the tab panels** so switching between the three tool tabs does not re-read or re-render the body.
- **Tool card** (one per tool tab), then the F-003 privacy notice, then the assurance line "Sem envio automático…" (on the three tool tabs).
- **Automático:** the F-013 contact selection (`ResponderPanel`) unchanged except the tab label; until F-013's UI ships, the static placeholder card.
- The F-003 "DIMSCALE AI Assistant" heading and the "Melhorar texto e tom" "Em breve" card are **removed** (architect proposal, Q-K3): the tabs replace the heading, and the tone options in Resposta deliver the useful part of F-006, so a card promising "melhorar texto e tom" would be misleading.

```
Resumo tab (auto mode switch shown only when AUTO_SUMMARY_AVAILABLE, F-003)
+--------------------------------------+
| DIMSCALE  Mail                       |  header <= 40 px, one h1
| Resumo Resposta Traduzir Automático  |  TabList small, one row
+--------------------------------------+
| Pedido de orcamento - Obra 203       |  message context (shared by
| Ana Silva <ana@exemplo.pt>           |  the 3 tool tabs)
| 02/01/2026 03:04                     |
| > Mensagem                           |  accordion, collapsed
| +- Resumo do email -------[c] [e]-+  |  h2 + Copiar, Expandir (icons,
| | Resumo automático         [ o] |  |  only when a result exists)
| | [ Resumir ]                     |  |  Switch, primary button
| | - ...                           |  |  result, max-height 240 px
| | - ...                           |  |  unless expanded
| | - ...                           |  |
| | [Pedido] [Prioridade alta] [Tom]|  |
| | Prazo: ...  Pedido: ...         |  |
| | Gerado por IA, pode estar errado.| |
| +---------------------------------+  |
| Ao usar a IA, ... enviados à OpenAI. |  F-003 privacy notice
| (shield) Sem envio automatico. ...   |  assurance
+--------------------------------------+
| (o) Ligado como Ana Silva            |  footer status
+--------------------------------------+

Resposta tab (card only; context, notice, assurance as above)
| +- Resposta sugerida -----[c] [e]-+  |
| | Idioma            Tom           |  |  two Select, wrap to two rows
| | [Idioma do e-mail v] [Normal v] |  |  when narrower than needed
| | [ Sugerir resposta ]            |  |
| | Bom dia Ana, ...                |  |  pre-wrap
| | Gerado por IA, pode estar errado.| |
| | [ Abrir resposta para revisão ] |  |  primary after a result
| +---------------------------------+  |

Traduzir tab
| +- Tradução --------------[c] [e]-+  |
| | Traduzir para [Português v]     |  |  Select
| | [ Traduzir ]                    |  |
| | Assunto: Quotation request ...  |  |  translated subject (if any)
| | Good morning, please find ...   |  |  pre-wrap body
| | Gerado por IA, pode estar errado.| |
| +---------------------------------+  |

Automático tab: the F-013 Responder UI (section "Pane: Responder tab" in F-013), label only renamed.
```

#### Card anatomy (`ToolCard`, one component for the three tool cards)
1. **Title row:** card title as `h2`; at the right, two Fluent subtle icon buttons shown **only when the card has a result**: **Copiar** (Copy icon, accessible name and tooltip "Copiar"; same clipboard logic and fallback text as F-003) and **Expandir** / **Reduzir** (expand icon, `aria-expanded`, toggles the result's max-height). No icon-only control without an accessible name.
2. **Options row** (per tool): Resumo the "Resumo automático" `Switch` (F-003); Resposta two `Select`s "Idioma" and "Tom"; Traduzir one `Select` "Traduzir para". Each `Select` has a visible `Label`. Options are kept in memory for the pane session (not reset by `ItemChanged`, reset when the pane reloads); not persisted (architect proposal, Q-K5).
3. **Primary button** (Resumir / Sugerir resposta / Traduzir) with the F-003 pending text and `aria-busy`; disabled while any AI request is pending (one at a time per pane, F-003).
4. **Result region:** text nodes only, `dir="auto"`, `overflow-wrap: anywhere`; collapsed state `max-height: 240px` with `overflow-y: auto`, `tabIndex=0` and an accessible name; expanded state has no max-height (the pane itself scrolls). "Expandir" is an in-pane expansion, not a dialog or a new window (architect proposal, Q-K2).
5. **Advisory** "Gerado por IA, pode estar errado." under each result; errors as a `MessageBar` inside the card (F-003 error table).
6. Resposta only: "Abrir resposta para revisão" (F-003, unchanged).

### Content set (production, 2026-10-23) (superseded 2026-10-29 by the section above)
```
+--------------------------------------+
| DIMSCALE  Mail                       |  header ≤ 40 px, one h1
| Analisar | Responder                 |  TabList
+--------------------------------------+
| Pedido de orcamento - Obra 203       |  Analisar: subject (wraps)
| Ana Silva <ana@exemplo.pt>           |  sender
| 02/01/2026 03:04                     |  date
| v Mensagem                           |  accordion, COLLAPSED by default
|   +--------------------------------+ |  when open: body preview, pre-wrap,
|   | Bom dia, segue o pedido...     | |  max-height ~120 px, focusable
|   +--------------------------------+ |  scroll region; (truncado) + length
| (sparkle) IA (resumo e resposta      |  one "em breve" card
|  sugerida) em breve.                 |
| (shield) Sem envio automatico.       |  assurance
+--------------------------------------+
| (o) Ligado como Ana Silva            |  footer status, 12 px
+--------------------------------------+
```
- **Header:** wordmark + "Mail"; one `h1` with accessible name "DIMSCALE Mail". No tagline.
- **Tabs:** "Analisar" (default) and "Responder", Fluent `TabList`. "Escrever" is not rendered (compose mode is F-010).
- **Analisar:** subject, sender (name and address, or address alone, never `<>`/"undefined") and date always visible; the body excerpt (F-001 limit 2,000 characters, "(truncado)" marker, length line) sits inside a Fluent `Accordion` item titled "Mensagem" that is **collapsed by default** (compact pane, orchestrator decision 2026-10-23; the body is still read on load so opening it is instant); the `comingSoon` card, the assurance line. States: loading (Skeleton, `aria-busy`), no message, body error with "Tentar novamente" (subject, sender and date still shown when the item has them).
- **Responder:** the F-013 contact selection UI ([F-013](F-013-automatic-reply-drafts.md), section "Pane: Responder tab"). Until that F-013 step ships, Responder shows one static card: "Respostas automáticas para contactos escolhidos: em breve." No fake controls.
- **Footer:** the F-002 status (table below), `role="status"`. Actionable states also show a Fluent `MessageBar` above the tab content.

### Backend status (F-002 logic unchanged; presentation only)
| State | Footer | Notice |
|---|---|---|
| connecting | "A ligar ao backend DIMSCALE…" | none |
| connected | "Ligado como {name}" (or "Ligado") | none |
| needSignIn | "Sessão não iniciada" | info: "Inicie sessão para ligar ao backend DIMSCALE" + "Iniciar sessão" (calls `signIn` only on click) |
| error | "Erro de ligação" (+ code) | error: the F-002 message + code + "Tentar novamente" |
| unsupported | "Sessão indisponível" | info: the F-002 unsupported text |
| notConfigured | "Backend não configurado" | info: the F-002 not-configured text |

### Ribbon buttons "IA" and "Resumo" (2026-10-29; replaces the single button below)
**Recommended approach (the smallest that relies only on documented behaviour):** two `ShowTaskpane` buttons in the same group, pointing at **two HTML pages built from the same template and the same JavaScript bundle**:

| Control id | Label | Supertip title / description (PT-PT) | `SourceLocation` resid → URL | Icons (16/32/80) |
|---|---|---|---|---|
| `msgReadOpenPaneButton` (kept, so the existing pinned state and any user customisation keep the same id) | **IA** | "DIMSCALE MAIL: IA" / "Abre o painel DIMSCALE Mail AI: resumo, resposta, tradução e respostas automáticas." | `Taskpane.Url` → `…/taskpane.html` (unchanged) | `Icon.Ia.16/32/80` → `assets/ia-16.png`, `ia-32.png`, `ia-80.png` |
| `msgReadSummaryButton` (new) | **Resumo** | "Resumo com IA" / "Abre o painel no separador Resumo e resume o e-mail aberto. O texto do e-mail é enviado à OpenAI." | `TaskpaneResumo.Url` → `…/resumo.html` | `Icon.Resumo.16/32/80` → `assets/resumo-16.png`, `resumo-32.png`, `resumo-80.png` |

- Both actions carry `<SupportsPinning>true</SupportsPinning>` (VersionOverrides v1.1, as AC-34). No `TaskpaneId`, no `Title` (both "Not supported in Outlook add-ins").
- The group stays `<OfficeTab id="TabDefault">` → one `<Group id="msgReadGroup">` labelled "DIMSCALE MAIL". The existing button **becomes "IA"** rather than a third button being added next to it (architect proposal, Q-K1): a "DIMSCALE MAIL" button and an "IA" button would both open the same pane, and Microsoft's guidance is "Don't add unnecessary buttons". This changes the 2026-10-23 answer to Q-R2 (button label "DIMSCALE MAIL"); the group label keeps the name.
- **Why a second page and not `taskpane.html?tab=resumo` or `#resumo`:** passing a query string or hash through to the page is not documented for Outlook and a community report says query parameters were stripped (facts table). A second page needs only a second `HtmlWebpackPlugin` output in `webpack.config.js` (same template, same `taskpane` chunk) and one check of `location.pathname` at start-up. Same origin, so NAA (redirect is origin-only), `ALLOWED_ORIGINS` and CORS are unchanged.
- **Entry behaviour:** the page records once, at load, whether `location.pathname` ends with `/resumo.html`. If it does, the pane opens on the **Resumo** tab and **starts one summary** for the message loaded at that moment, if the F-003 click preconditions hold (message loaded, non-empty body, silent token). A click on a button labelled "Resumo", whose tooltip says the text is sent to OpenAI, is the user's action, so this is not an auto-run. Later `ItemChanged` events do **not** start a summary from the entry (only auto mode does that, F-003). From `taskpane.html` the pane opens on Resumo and the entry starts nothing (F-003 auto mode, if the user switched it on, applies on either page). Query string and hash are never read.
- **What is undocumented and must be checked live before the admin upload (AC-52):** with the pane already open (pinned or not), does clicking the other button (a) reload the pane at the other page, (b) leave the open pane as it is, or (c) open a second pane? Does pinning survive it? If any client shows (c) or loses pinning, the fallback is that both buttons point at `taskpane.html` (Resumo then only opens the pane) or the Resumo button is removed; the user decides (AC-53). Design rule: nothing in the pane depends on (a) happening.
- **Reload and the auto-mode preference:** Microsoft documents that `roamingSettings` resets to the values read at first load when the pane "navigates to another page or reloads". If (a) happens, a user who switched auto mode off in that pane session could see it back on after the reload. AC-52 records whether this happens; F-003 handles it (F-003 AC-50).
- **Icons:** new SVG masters `assets/ia-master.svg` (a sparkle, the usual AI metaphor) and `assets/resumo-master.svg` (a page with text lines and a small sparkle modifier at the bottom right), Monoline style, no text characters, transparent background, the DIMSCALE green of `icon-master.svg` as the one accent colour on a dark-grey outline, PNG 32-bit RGBA at 16, 32 and 80 px. Legible on light and dark ribbons at 16 and 32 px (person-checked). The add-in icons `icon-64.png`/`icon-128.png` (`IconUrl`, `HighResolutionIconUrl`) and the existing `icon-16/32/80.png` files stay (no deletion).
- **Version:** `manifest.xml` `<Version>` `1.0.1.0` → **`1.0.2.0`**; `manifest.customtab-test.xml` gets the same two buttons inside its custom tab with unique control ids and its own `<Version>` **`1.0.2.1`** (still its own `<Id>`, never deployed).
- **Admin:** this manifest is the one sent to the tenant admin for centralized deployment. Nothing goes to the admin before AC-52 is recorded and the user confirms (AC-53). Any later manifest change needs a new `<Version>` and another admin re-upload (F-011 runbook).

### Ribbon entry "DIMSCALE MAIL" (manifest) (single-button version of 2026-10-23, superseded 2026-10-29 by the section above; the group, `TabDefault`, pinning and custom-tab rules still apply)
- **v1 approach (documented on every platform):** `MessageReadCommandSurface` → `<OfficeTab id="TabDefault">` → **one** `<Group>` with label "DIMSCALE MAIL" → one `Button` whose label is also **"DIMSCALE MAIL"**, `Supertip` title "DIMSCALE MAIL", `Action xsi:type="ShowTaskpane"` with `<SupportsPinning>true</SupportsPinning>`. In classic Outlook this shows a "DIMSCALE MAIL" group on the default (Home/Base) tab; in new Outlook and the web the command sits on the default tab for that platform and mode, where the simplified ribbon may show only the icon (documented for Office on the web), so the button label and tooltip carry the name.
- **No `CustomTab` in the production manifest until the experiment below passes** (user wants an own tab, 2026-10-23). Microsoft documents `CustomTab` as classic-Outlook-on-Windows only and does not document what the other clients do with a manifest that declares it, nor whether Outlook accepts `OfficeTab` and `CustomTab` together.
- No group `Icon` (not supported in Outlook). Button icons 16/32/80 as today.
- Button label: **"DIMSCALE MAIL"** (user decision 2026-10-23, Q-R2 closed). The working `manifest.xml` still has the label "Abrir", which the coder replaces.

### Custom tab experiment (user wants a "DIMSCALE MAIL" tab; 2026-10-23)
What the docs say: the tab is **not available** in Outlook on the web, on Mac and in new Outlook on Windows ("you can put *custom* groups of controls on one of the built-in OfficeTabs instead"); only classic Outlook on Windows supports it. So on Mac and new Outlook a custom tab is **documented as unavailable**; the open point is only whether those clients ignore it gracefully (and still show the `TabDefault` group) or reject the manifest. The user's clients are classic Windows, new Outlook and Mac, so a tab that works only in classic is a partial result; the production manifest adopts it only if all three behave.
- **Test manifest** `manifest.customtab-test.xml` (repo root, never deployed, not built into `dist/` by default): a copy of the production manifest with a **different `<Id>`** (new GUID), `DisplayName` "DIMSCALE MAIL (teste)" and a distinct `<Version>`. Variant **B** (try first): one `MessageReadCommandSurface` extension point with `<CustomTab id="DIMSCALE.TabMail">` (its own `Group`, the same `Button`, tab `Label` "DIMSCALE MAIL") **plus** the existing `<OfficeTab id="TabDefault">` group with a distinct group ID and control IDs. If B fails validation or install, variant **A**: `CustomTab` only (no `TabDefault`), same file name with suffix `-a`. Same task pane URL, same permission `ReadItem`, `SupportsPinning` kept.
- **Procedure:** the user sideloads the test manifest in classic Windows, new Outlook and Mac (procedure per client: https://learn.microsoft.com/office/dev/add-ins/outlook/testing-and-tips, not verified per client here) and records, per client: install result or error text, where "DIMSCALE MAIL" appears (own tab, default tab, overflow, nowhere), whether the button opens the pane, whether pinning works. Then the user removes the test add-in.
- **Decision rule:** all three clients install it and each shows a usable "DIMSCALE MAIL" entry (own tab in classic is enough; the default-tab group is acceptable on the others) and no client shows two competing entries that confuse users, then the production manifest adopts variant B (a spec amendment and a new approval by the user first). Otherwise production keeps the `TabDefault` group only.
- Any manifest change ships with the centralized deployment (F-011 runbook: `<Version>` bump, admin upload). After that, code-only changes reach users without the admin.

### Pinning and item change
- `SupportsPinning` as above (VersionOverrides v1.0 and v1.1 both present).
- **Behaviour (user, 2026-10-23):** once pinned, the pane stays open across messages until the user closes it. The "DIMSCALE MAIL" ribbon button reopens it. Whether it reopens by itself after an Outlook restart is **undocumented**, so it is a live check (AC-41); the design must not depend on it.
- Register `ItemChanged` with `Office.context.mailbox.addHandlerAsync` in `Office.onReady`. On change: discard any pending body read for the previous item (its result must not be shown), reset the Analisar state, read the new item; if `item` is null show the "no message" state. The Responder tab and the F-002 status are not reset by an item change. (2026-10-29: "Analisar state" = the message context and the three tool cards' results, errors and pending state; the selected tab, the card options and the Automático tab are not reset.)

## Strings (PT-PT, all in `src/taskpane/strings.ts`)
F-001/F-002 strings as approved on 2026-10-22 (D6 = PT-PT), plus: `title` "DIMSCALE Mail"; tabs "Analisar", "Responder"; `comingSoon` "IA (resumo e resposta sugerida) em breve."; `assurance` "Sem envio automático. O envio é sempre feito por si."; `responderSoon` "Respostas automáticas para contactos escolhidos: em breve."; footer labels from the status table; accessible names "Pré-visualização do corpo", "A carregar a mensagem…". European Portuguese wording ("iniciar sessão", "ligar", "e-mail"). `html lang="pt-PT"`. Manifest ribbon strings "DIMSCALE MAIL" are a product name and stay as is.

**Added 2026-10-29 (PT-PT, for the user's review; the AI card strings are listed in the F-003 plan):**
| Key (suggested) | Text |
|---|---|
| `tabSummary` / `tabReply` / `tabTranslate` / `tabAuto` | Resumo / Resposta / Traduzir / Automático (replace "Analisar" and "Responder") |
| `aiExpand` / `aiCollapse` | Expandir / Reduzir (icon button names and tooltips) |
| `aiCopy` (existing) | Copiar (now also the icon button name and tooltip) |
| `aiTranslationTitle` | Tradução |

Manifest strings (in `manifest.xml` `Resources`, PT-PT, within the 125/250-character limits): `GroupLabel` "DIMSCALE MAIL" (unchanged); `TaskpaneButton.Label` "IA"; `IaButton.Title` "DIMSCALE MAIL: IA"; `TaskpaneButton.Tooltip` "Abre o painel DIMSCALE Mail AI: resumo, resposta, tradução e respostas automáticas."; `ResumoButton.Label` "Resumo"; `ResumoButton.Title` "Resumo com IA"; `ResumoButton.Tooltip` "Abre o painel no separador Resumo e resume o e-mail aberto. O texto do e-mail é enviado à OpenAI."

## Acceptance criteria
"Mocks" = vitest/jsdom with mocked Office.js. "Visual" = a person in a browser. "Live" = Outlook (classic on Windows Version 2409+ unless stated), build recorded.

### Preserved behavior
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-1 | Every F-001 AC (1–10) and every pane-related F-002 AC (1–7, 16, 17, 18, 20) still holds, with the PT-PT literals approved on 2026-10-22 replacing the English ones. No assertion removed or weakened; the vitest count does not drop below the step-0 baseline; `deno test` unchanged. | Mocks |
| F-012/AC-2 (amended 2026-10-29) | The pane reads only `subject`, `from`, `dateTimeCreated`, `body` from the item (Proxy test kept). The only new Office.js calls are the theme calls (AC-7), `ItemChanged` registration (AC-35), `requirements.isSetSupported`, the F-003 reply-form calls and `Office.context.roamingSettings.get/set/saveAsync` for the one F-003 auto-mode key. Add-in permission stays `ReadItem`. (The 2026-10-23 clause "no new network request; no change under `supabase/`, `auth.ts`, `backend.ts`, `config.ts`" is superseded by F-003, which adds `analyze-message`.) | Mocks, diff |
| F-012/AC-3 | Email-derived and name text are React text nodes; no `dangerouslySetInnerHTML`, `innerHTML`, `insertAdjacentHTML` or `href` from email content; F-001 injection tests pass. | Mocks, grep |
| F-012/AC-4 | No `console.*` in new code; no `localStorage`, `sessionStorage`, cookies or IndexedDB. | Grep |

### Theme
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-5 (amended) | The pane uses `webLightTheme`/`webDarkTheme` (and the Fluent high-contrast theme under forced colours). No custom brand ramp, no DIMSCALE accent colour in the theme. | Mocks |
| F-012/AC-6 (amended) | The Office body-colour guard: a test feeds pairs above and below 4.5:1 and asserts Office colours are applied only above it; otherwise the Fluent neutrals are used. Semantic pairs actually rendered (MessageBar intents, status icons) are measured in both themes and listed in the plan evidence. | Mocks, Visual |
| F-012/AC-7 | Theme resolution as in "Theme": (a) 1.14 + dark `bodyBackgroundColor` → dark; (b) light → light; (c) no 1.14 + `prefers-color-scheme: dark` → dark; (d) neither → light; (e) invalid hex → step 2; (f) `forced-colors` → high contrast; (g) `OfficeThemeChanged` registered via `addHandlerAsync`, applied when fired, removed on unmount; (h) a mock without `officeTheme` or `isSetSupported` does not throw; (i) `isDarkTheme`/`themeId` absent from `src/`. | Mocks |
| F-012/AC-8 | Live: for each Office theme offered (record the list), the pane follows it with no seam the user considers a defect; record `bodyBackgroundColor` per theme and whether a change with the pane open applies live. Repeat in new Outlook and on Mac (the user's clients, D2; Mac see AC-42). | Live (user) |
| F-012/AC-9 | Live: with the four Windows contrast themes, all text, buttons, focus and status icons are visible; record whether `forced-colors` applies. | Live (user) |

### Layout and content
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-10 (amended; **superseded 2026-10-29 by AC-43 and AC-45**) | Production renders exactly the content set: header, tabs, Analisar (subject, sender and date, the "Mensagem" accordion **collapsed on open**, `comingSoon`, assurance), Responder (placeholder card until F-013's UI step, then the F-013 UI), footer, notices when needed. No other section, no disabled fake control. | Mocks |
| F-012/AC-11 (amended) | Header is one row ≤ 40 px at 320 px: inline wordmark SVG with an accessible name, the text "Mail", exactly one `h1` named "DIMSCALE Mail"; wordmark black on light, white on dark. | Mocks, Visual |
| F-012/AC-12 (amended) | Message context shows subject, sender (rules above) and date always. The body excerpt is inside a Fluent `Accordion` item "Mensagem" that is **collapsed by default** (the body text is not in the accessibility tree or tab order until the user opens it; opening it needs a click or Enter/Space, never happens on its own, and an item change collapses it again). When opened it shows the excerpt with "(truncado)" when cut and the length line. All F-001 edge cases hold (empty subject/body, no item, loading, body error with Retry and `role="alert"`, large body, emoji, injection text); the body error notice and "Tentar novamente" stay visible **outside** the collapsed accordion. On body error, subject, sender and date are still shown. Body text ≤ 2,000 characters, whitespace kept. F-001 ACs that check the body text are verified after opening the accordion in the test. | Mocks |
| F-012/AC-13 | Backend status renders the six F-002 states per the table, same state machine and handlers, visible text for every state, footer `role="status"`, error notice `role="alert"`, sign-in popup only from the "Iniciar sessão" click. | Mocks |
| F-012/AC-14 | **Superseded** (future-slot components dropped by the compact content set). | — |
| F-012/AC-15 (amended 2026-10-29) | The assurance line is present in every state of the Resumo, Resposta and Traduzir tabs that shows the message (formerly "every Analisar state"), static, not removable by a prop. | Mocks |
| F-012/AC-16 (amended; **superseded 2026-10-29 by AC-43**) | Tabs: `TabList` semantics (`role="tablist"`, `aria-selected`, arrow keys), exactly "Analisar" and "Responder", Analisar selected on open; "Escrever" cannot render. Switching tabs keeps the message state. | Mocks |

### Responsive and accessibility
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-17 | At 320, 360, 450 and 500 px with a 200-character unbroken subject, a 120-character address and a 2,000-character body: no horizontal scroll, nothing clipped or overlapping. | Visual, then Live at the default and a narrow and wide pane |
| F-012/AC-18 | Reflow at 320 CSS px (WCAG 1.4.10) and usable at 200% text (1.4.4). | Visual, Live |
| F-012/AC-19 | Keyboard: all controls reachable with Tab in visual order, tabs with arrows, Enter/Space activate, preview scroll region focusable, visible focus ≥ 3:1, no trap, focus never moved unless the user acts (including on `ItemChanged`). | Mocks, Live |
| F-012/AC-20 | Semantics: landmarks, one `h1`, named regions, decorative icons `aria-hidden`, `lang="pt-PT"`, `aria-busy` while loading, `role="status"`/`role="alert"`. Narrator read-out recorded per state. | Mocks, Live |
| F-012/AC-21 | No state by colour alone; status icons differ in shape. | Mocks, Visual (grayscale) |
| F-012/AC-22 | No custom animation; nothing of ours runs under `prefers-reduced-motion: reduce`; Fluent Skeleton behaviour recorded. | Grep, Visual |

### Strings, icons, manifest, build
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-23 | Every user-visible string comes from `strings.ts` (allow-list test or grep); PT-PT reviewed by the user. | Mocks, user review |
| F-012/AC-24 (amended) | `assets/icon-16/32/64/80/128.png`: exact sizes, 32-bit RGBA, the green "D" on white chosen by the user, SVG master in the repo, legible at 16 and 32 px on light and dark ribbons (person-checked). | Mocks (PNG header), Visual |
| F-012/AC-25 | Live: ribbon and File > Manage add-ins show the new icons; record what refreshed a cached icon. | Live |
| F-012/AC-26 (amended; **superseded 2026-10-29 by AC-48** for the 2026-10-29 manifest change) | Manifest diff contains only: the ribbon entry of AC-32 (button label "DIMSCALE MAIL" instead of "Abrir"), `SupportsPinning` (AC-34), the icon files if renamed, and a `<Version>` bump. Permission stays `ReadItem`; no `CustomTab` in the production manifest (until AC-40 is decided and the user approves an amendment); `npm run validate` and `manifest.test.ts` pass; `dist/manifest.xml` has the production URLs. | Diff, command output |
| F-012/AC-27 | No new network host (no webfont, CDN, external image, analytics). F-011 checks still pass (hashed names, no source maps, no secrets, no `localhost` in the production manifest). Bundle size delta recorded against the step-0 budget. | Build output, grep |
| F-012/AC-28 | Dependency change limited to `@fluentui/react-icons` as a direct dependency at the version already in the lockfile; `npm audit` recorded with no new High/Critical. | Diff, command output |
| F-012/AC-29 | Typecheck, lint, tests, `npm run build`, `build:dev`, `validate` pass. After merge, the F-011 workflow deploys. | Command output |
| F-012/AC-30 (amended) | Visual acceptance by the user in classic Outlook (then new Outlook and Mac), light and dark, default/narrow/wide pane, for: message loaded (accordion collapsed and open), loading, no message, body error, the reachable backend states. Screenshots without real email content. | Live (user) |
| F-012/AC-31 | **Superseded** (the prototype gate was met: direction D chosen 2026-10-23). | — |

### Ribbon and pinning (new)
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-32 (**superseded 2026-10-29 by AC-48**: two buttons) | The manifest's `MessageReadCommandSurface` has exactly one `<OfficeTab id="TabDefault">` with exactly one `<Group>` labelled "DIMSCALE MAIL" containing one `Button` labelled "DIMSCALE MAIL" (Supertip title "DIMSCALE MAIL") whose action is `ShowTaskpane`. No `<CustomTab>` in `manifest.xml` (the test manifests of AC-37 are separate files). | `manifest.test.ts`, `npm run validate` |
| F-012/AC-33 (amended) | Live: after the add-in is installed from the production manifest, **classic Outlook on Windows, new Outlook on Windows and Outlook on Mac** each show a "DIMSCALE MAIL" entry and clicking it opens the pane. Record per client where it appears and what text or icon is visible. | Live (user) |
| F-012/AC-34 | `<SupportsPinning>true</SupportsPinning>` is inside that button's `ShowTaskpane` action, in the VersionOverrides v1.1 block nested in v1.0. | `manifest.test.ts`, `npm run validate` |
| F-012/AC-35 | `ItemChanged`: registered once via `addHandlerAsync` after `Office.onReady`; on fire the Analisar view shows the new item; a body read started for the previous item that resolves late is discarded (test with two items and a delayed first `getAsync`); `item === null` gives the "no message" state; the active tab, the Responder state and the F-002 status are unchanged; focus is not moved. | Mocks |
| F-012/AC-36 (amended) | Live, in **classic Windows, new Outlook and Mac** (Mac 16.13 or later; record versions): pin the pane, select three different messages in turn, then a calendar item or nothing: the pane follows each message with no stale content and no error. Record the client and build. | Live (user) |

### Custom tab experiment, theme and pinning on Mac (new, 2026-10-23)
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-37 | `manifest.customtab-test.xml` exists as a separate file with its own `<Id>` (different from production), a distinct `<Version>` and DisplayName "DIMSCALE MAIL (teste)". Variant B contains one `<CustomTab>` (label "DIMSCALE MAIL", resid ≤ 32 chars) and one `<OfficeTab id="TabDefault">` with unique group and control IDs, in `MessageReadCommandSurface`; permission `ReadItem`; no other change from production. It passes `npm run validate`. `manifest.xml` and the CI deploy are untouched, and the test manifest is not published to Pages. | `manifest.test.ts` (or a script), `npm run validate`, diff |
| F-012/AC-38 | Live, per client (classic Windows, new Outlook, Mac), the user records for the test manifest: installs or error text, where "DIMSCALE MAIL" shows, whether the button opens the pane, whether pinning works. If variant B fails validation or install anywhere, the same record is made for variant A. | Live (user), table in the plan |
| F-012/AC-39 | Results are written in the plan with date, client, build/version and screenshots without email content. Microsoft's statement that `CustomTab` is unavailable on Mac and new Outlook is compared with what was observed; any difference is recorded as a finding. | Live (user), review |
| F-012/AC-40 | Decision: the production manifest adopts the custom tab only if all three clients behave per the decision rule in "Custom tab experiment". Otherwise it stays with the `TabDefault` group. Adoption needs a spec amendment approved by the user, a `<Version>` bump and an admin re-upload. | User decision |
| F-012/AC-41 | Live: with the pane pinned, selecting other messages leaves it open until the user closes it; after closing, the "DIMSCALE MAIL" button reopens it; after quitting and restarting Outlook, record whether the pane reopens by itself (undocumented). Per client: classic Windows, new Outlook, Mac. | Live (user) |
| F-012/AC-42 | Live on Mac: record the Outlook version and UI (classic or new), `isSetSupported("Mailbox","1.14")` and `isSetSupported("NestedAppAuth","1.1")` results (shown by a temporary diagnostic or the dev console, not committed), whether the pane follows the Office theme, and whether `OfficeThemeChanged` fires when the theme changes with the pane open. Expected from the docs: new Mac UI has 1.14, classic Mac UI (max 1.8) falls back to `prefers-color-scheme`. | Live (user) |

### Kutools-style pane and ribbon buttons (new, 2026-10-29)
The AI behaviour inside the cards (requests, auto mode, options, errors) is tested under F-003 (AC-35 onward); these ACs cover layout, entry and manifest.

| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-43 | Tabs: `TabList` semantics (`role="tablist"`, `aria-selected`, arrow keys), exactly **Resumo, Resposta, Traduzir, Automático** in that order, labels from `strings.ts`; Resumo selected on open from either page; "Analisar", "Responder" and "Escrever" never render. Switching tabs keeps the message state and each card's result, options and pending state, and makes no request by itself (F-003 auto mode is the only exception, F-003 AC-43). | Mocks |
| F-012/AC-44 | At 320, 360, 450 and 500 px the tab row is one line with all four labels fully visible, no horizontal scroll, no clipped focus outline; at 200 % text it may wrap but nothing is clipped or overlapping. | Visual, then Live |
| F-012/AC-45 | Message context (subject, sender, date, "Mensagem" accordion collapsed on open, F-001/F-012 states and edge cases of AC-12) renders once above the tab panels on Resumo, Resposta and Traduzir and is absent on Automático; switching between the three tool tabs does not call `body.getAsync` again (call count asserted). The F-003 privacy notice and the assurance line show on the three tool tabs. The "DIMSCALE AI Assistant" heading and the "Melhorar texto e tom" card are gone. | Mocks |
| F-012/AC-46 | Card anatomy: each tool card has an `h2` title, its options row, its primary button, a result region and the advisory line. "Copiar" and "Expandir" icon buttons appear only when the card has a result, each with an accessible name and tooltip from `strings.ts`; "Expandir" sets `aria-expanded="true"`, removes the 240 px max-height and changes its name to "Reduzir"; collapsed, the result region is a focusable scroll region with an accessible name. "Copiar" uses the F-003 clipboard logic and fallback. Keyboard: all of them reachable with Tab and operable with Enter/Space. | Mocks |
| F-012/AC-47 | Automático = the F-013 `ResponderPanel` with only the tab label changed; every F-013 pane test passes unchanged except the tab-label assertion (replaced, not removed); its state is not reset by tab switches or `ItemChanged`. | Mocks |
| F-012/AC-48 | `manifest.xml`: one `<OfficeTab id="TabDefault">` with one `<Group id="msgReadGroup">` labelled "DIMSCALE MAIL" containing exactly two `Button` controls: `msgReadOpenPaneButton` (label "IA") → `ShowTaskpane` `Taskpane.Url` = `…/taskpane.html`, and `msgReadSummaryButton` (label "Resumo") → `ShowTaskpane` `TaskpaneResumo.Url` = `…/resumo.html`; both with `<SupportsPinning>true</SupportsPinning>` in the v1.1 block; each with its own 16/32/80 icons and the supertips of "Strings"; no `TaskpaneId`, no `Title`, no `<CustomTab>`; permission `ReadItem`; `<Version>` `1.0.2.0`; every resid ≤ 32 characters, labels ≤ 125, descriptions ≤ 250. `npm run validate` and `manifest.test.ts` pass; `dist/manifest.xml` has the Pages URLs for both pages and all six icons. | `manifest.test.ts`, `npm run validate`, diff |
| F-012/AC-49 | `manifest.customtab-test.xml` gets the same two buttons inside its existing custom tab, with control ids unique in that file, `<Version>` `1.0.2.1`, its own `<Id>` unchanged; it passes `npm run validate` and is not published to Pages. (The file in the repo is variant A, `CustomTab` only, DisplayName "DIMSCALE Mail AI (teste separador)", while AC-37 describes variant B and "DIMSCALE MAIL (teste)"; this amendment does not change that, see Q-K6.) | `manifest.test.ts` or a script, `npm run validate` |
| F-012/AC-50 | Entry page: `npm run build` emits `resumo.html` from the same template with the same `taskpane` bundle (no new chunk, no inline script). At start-up the pane reads `location.pathname` once: ending in `/resumo.html` → Resumo tab and one F-003 summary start for the item loaded at that time if the click preconditions hold (exactly one `analyze-message` request; none when the body is empty, the token needs interaction or the item is null); any other path → Resumo tab and zero requests from the entry (auto mode, if on, behaves per F-003 AC-43). A later `ItemChanged` never starts a summary because of the entry page. `location.search` and `location.hash` are not read anywhere in `src/` (grep). | Mocks (stubbed `location.pathname`), build output, grep |
| F-012/AC-51 | `assets/ia-16.png`, `ia-32.png`, `ia-80.png`, `resumo-16.png`, `resumo-32.png`, `resumo-80.png`: exact sizes, 32-bit RGBA with transparent pixels, made from the committed `assets/ia-master.svg` and `assets/resumo-master.svg`; no text characters in the artwork; legible at 16 and 32 px on light and dark ribbons (person-checked). | `icons.test.ts` (PNG header and alpha), Visual |
| F-012/AC-52 | Live, **before anything is sent to the tenant admin**, with the new manifest sideloaded and the pane (with `resumo.html`) served from Pages, in classic Windows, new Outlook and Mac (record versions): (1) where both buttons appear and what label or icon is visible; (2) "IA" with the pane closed opens it on Resumo with no AI request (auto mode off); (3) "Resumo" with the pane closed opens it on Resumo and one summary starts; (4) with the pane open and pinned from "IA", clicking "Resumo": the pane reloads at `resumo.html` (summary starts), stays unchanged, or a second pane opens; and the reverse; (5) whether pinning survives (4); (6) with auto mode on at pane open, switched off, then (4): whether auto mode comes back on after a reload (F-003 AC-50). Test emails only (D5). | Live (user), table in the plan |
| F-012/AC-53 | Decision before the admin upload: if every client shows a usable entry for both buttons and none opens a second pane or loses pinning, the manifest of AC-48 is sent to the admin (F-011 runbook, `<Version>` 1.0.2.0). Otherwise the user chooses: both buttons on `taskpane.html` (Resumo only opens the pane), or remove "Resumo"; either needs a spec amendment, a new `<Version>` and the AC-52 recheck. The admin gets one final manifest, once. | User decision, recorded in `status.md` |
| F-012/AC-54 | Visual acceptance by the user (extends AC-30) at 320 px, light and dark, of: each tool tab with no result, pending, result collapsed and expanded, error; Automático; screenshots without real email content. | Live (user) |

## Edge cases
- Very long unbroken subject, address or body line: wraps, no horizontal scroll.
- Sender without name: address only; both empty: "De" row with an em dash.
- RTL or mixed text: `dir="auto"`.
- Rapid item switching while pinned: only the last item's body is shown (AC-35).
- Pinned pane on a calendar item: pinning is not available for appointments (Microsoft); the pane may close or show "no message"; record (AC-36).
- Office returns a body colour pair with poor contrast: Fluent neutrals are used (AC-6).
- Client without Mailbox 1.14: `prefers-color-scheme`, then light. Such clients cannot sign in anyway (F-002).
- Backend slow or down: message context works independently of the footer.
- Simplified ribbon on the web: only the icon may show (documented); the tooltip carries "DIMSCALE MAIL".
- High contrast: Fluent high-contrast theme.
- (2026-10-29) User clicks "Resumo" while the pane is already open: undocumented (AC-52); the pane must work whether it reloads or not.
- (2026-10-29) `resumo.html` opened with no message selected or on a calendar item: no summary starts, the "no message" state shows.
- (2026-10-29) The pane is restored by Outlook from `resumo.html` without a click (for example after a restart, undocumented, AC-41): one summary would start. Bounded by the caps; recorded in AC-41/AC-52; if it happens the user decides whether the entry should only select the tab.
- (2026-10-29) Old manifest (1.0.1.0) still installed while the new pane is deployed: the single button opens `taskpane.html`, which works unchanged; `resumo.html` exists but is not referenced. Code and manifest can therefore ship in either order.

## Out of scope
- Any AI feature or backend call (F-003 onward). The F-013 contact UI itself is specified and built under F-013.
- `CustomTab` in the production manifest (only the experiment ACs 37 to 40, then a separate approval), compose mode and "Escrever" (F-010), localization machinery, first-run experience, user theme override, persisted preferences, webfonts, analytics.
- Changing `auth.ts`, `backend.ts`, `config.ts` or Supabase functions.
- (2026-10-29, user decision 6) model selector, languages other than Portuguese and English, "Email Learning" (F-016), reading-pane or single-page-view toggles; the Kutools floating edit button (compose is F-010); expanding a result into a dialog or separate window; persisting the reply/translate options; a third ribbon button.

## Risks
- **CustomTab behaviour unknown outside classic Outlook** (documented as unavailable on Mac, web and new Outlook). Production uses only `TabDefault`; the test manifest (AC-37) is separate with another `<Id>`, so a bad result can't break the production install. Two add-ins with different IDs sideloaded together on one client may both show; the user removes the test one afterwards.
- **Classic Mac UI has Mailbox 1.8 only**: no theme API and no 1.14 features there; the pane falls back to `prefers-color-scheme` (AC-42).
- **Theme detection under-documented for Outlook** (values per theme, whether the event fires in classic): safe fallback; AC-8 live.
- **Stale content in a pinned pane** if a late body read wins: AC-35 test.
- **Manifest change needs the admin**: bundled with the planned centralized deployment.
- **Existing tests assert exact strings and DOM shapes**: the main mechanical risk to AC-1.
- **(2026-10-29) Two `ShowTaskpane` buttons in Outlook are undocumented** (`TaskpaneId` not supported, no shared runtime). Mitigation: two pages with no dependency on reload behaviour, live check AC-52 before the admin upload, fallback AC-53.
- **(2026-10-29) Four tabs at 320 px** may not fit in one row (AC-44); the user picks a shorter label rather than the coder (Q-K4).
- **(2026-10-29) Two "automatic" concepts** (the Automático tab = automatic drafts; the "Resumo automático" switch = auto summaries) may confuse users. Mitigation: distinct labels and the switch's own privacy dialog; user review of strings.
- **(2026-10-29) Manifest re-upload:** the admin has not received a manifest yet, so this change costs nothing extra if it lands before the first upload (AC-53).

## Brandbook rules (from the DIMSCALE Brand Book, 2026-10-22; source files stay in the git-ignored `Modelo_DIMSCALE/`)
- The logo is used only in black or white, never coloured. Black wordmark on light, white on dark.
- Digital wordmark minimum about 20 mm (~76 px); clear space = logo height.
- For e-mail and online platforms the brandbook prescribes Helvetica; the pane uses Segoe UI (Outlook-native, direction D) for text and the wordmark as SVG, so no font is bundled.

## User decisions
- 2026-10-22: UI language **PT-PT** (F-001/F-002 literals amended); icon **green "D" on white** (icon-only deviation from the black/white logo rule; exact green from the prototype); "Competências" = "Resumo da conversa" (later feature); pane easy to open/close → **pinnable task pane** (no left docking exists for web add-ins: https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane).
- 2026-10-23: **direction D**, **320 px compact**, **"DIMSCALE MAIL" ribbon entry**, **Responder = contact selection for automatic drafts (F-013)**.
- 2026-10-29 (from Kutools AI screenshots, relayed by the orchestrator): tabs **Resumo | Resposta | Traduzir | Automático**, each tool a card with copy and expand; auto summary per user, off by default (F-003); translate pt/en (F-003); reply language match/pt/en and tone (F-003); ribbon buttons **"IA"** and **"Resumo"** in "DIMSCALE MAIL", done before the admin upload; out of scope: model selector, other languages, Email Learning (F-016), reading-pane toggles.

## Open questions
- **Q-R1 (custom tab): answered 2026-10-23, the user wants an own "DIMSCALE MAIL" tab.** Handled by the experiment ACs 37 to 40. Still open: does new Outlook or Mac reject, ignore or fall back when the manifest declares `CustomTab` (undocumented), and does Outlook accept `CustomTab` with `OfficeTab` in one extension point (undocumented; variant A is the fallback)?
- **Q-R2: answered 2026-10-23, "DIMSCALE MAIL".** Closed.
- **Q-R3.** How does the user sideload a test manifest on Mac (and in new Outlook)? The testing-and-tips page is cited but not checked per client.
- **Q-R4.** Does a pinned pane reopen after an Outlook restart (AC-41)? Undocumented.
- **Q-R5.** Does `OfficeThemeChanged` fire on new Mac UI (AC-42)? Requirement set says supported, firing not documented.
- **Q11 (kept).** Does Outlook show the add-in name above the pane? If yes, the header may keep only the wordmark.
- **Q12 (kept).** Office theme values per theme, `OfficeThemeChanged` in classic, `forced-colors` in the webview: answered by AC-8, AC-9.
- **Q17 (kept).** Real pane widths in classic Outlook: measure live.

### Added 2026-10-29 (architect proposals that need the user's confirmation)
- **Q-K1.** The existing "DIMSCALE MAIL" button becomes "IA" (two buttons in the group), instead of keeping it and adding "IA" next to it (three buttons, two doing the same). Confirm.
- **Q-K2.** "Expandir" grows the card inside the pane (no dialog or separate window). Enough, or do you want a larger window like Kutools (an Office dialog, more code and its own runtime)?
- **Q-K3.** Remove the "Melhorar texto e tom" "Em breve" card now that Resposta has tone options. Confirm.
- **Q-K4.** If "Resumo Resposta Traduzir Automático" doesn't fit at 320 px, which shorter label: "Auto" for "Automático"?
- **Q-K5.** Reply/translate options are not remembered between pane sessions. Remember them too (same `roamingSettings` mechanism as auto mode, a few lines)?
- **Q-K6.** The test manifest in the repo is variant A ("DIMSCALE Mail AI (teste separador)", `CustomTab` only), not the variant B / "DIMSCALE MAIL (teste)" of AC-37. Was the custom-tab experiment (AC-38) run, and with which variant? Its result decides whether the final manifest has the own tab, so it should be settled before AC-53.
