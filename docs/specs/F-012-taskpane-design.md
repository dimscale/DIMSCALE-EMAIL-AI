# F-012: Task pane design (Outlook-native, 320 px compact, "DIMSCALE MAIL" ribbon entry)

- **Status:** Approved (2026-10-23) with the user's choices: **direction D "Outlook nativo"** (Fluent stock themes plus the user's Office theme colours; only the wordmark and the name are DIMSCALE), **320 px compact** (only the information the user needs), and a **"DIMSCALE MAIL" ribbon entry** that opens a pinnable pane. The "Responder" tab hosts the contact selection for automatic drafts (F-013). Earlier directions A, B and C are dropped.
- **Amended 2026-10-23** (requirements-architect) to match that approval. Superseded and removed: directions A/B/C, the DIMSCALE navy/teal Fluent ramps, the brand contrast table, the full-layout wireframe with future AI slots, the future-slot components, and "no manifest change". AC IDs are kept; superseded ACs are marked.
- **Amended again 2026-10-23** (requirements-architect) with the user's answers and one orchestrator decision: button label "DIMSCALE MAIL" (Q-R2 closed); the user wants an own "DIMSCALE MAIL" tab (Q-R1 becomes the experiment ACs 37 to 40, production manifest unchanged until it passes); clients in use are classic Windows, new Outlook and Mac (D2: Mac facts and live checks added); pinned-pane behaviour clarified (AC-36, AC-41); the message excerpt is **collapsed by default** in a "Mensagem" accordion (AC-10, AC-12).
- **Tier:** Active item, presentation, plus the manifest (ribbon entry, pinning). No Graph, no new add-in permission (`ReadItem` stays), no new backend call, no AI.
- **Depends on:** F-001, F-002, F-011 (all Done).
- **Plan:** [docs/plans/F-012-taskpane-design.md](../plans/F-012-taskpane-design.md)

## Goal
A compact pane that looks like part of Outlook, opened from a "DIMSCALE MAIL" ribbon entry, pinnable, that:
1. keeps everything F-001 and F-002 do (message context, backend status);
2. follows the user's Outlook theme (light/dark) at 320 px;
3. has two tabs: **Analisar** (the open message) and **Responder** (contact selection for automatic drafts, filled by F-013).

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

### Content set (production)
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

### Ribbon entry "DIMSCALE MAIL" (manifest)
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
- Register `ItemChanged` with `Office.context.mailbox.addHandlerAsync` in `Office.onReady`. On change: discard any pending body read for the previous item (its result must not be shown), reset the Analisar state, read the new item; if `item` is null show the "no message" state. The Responder tab and the F-002 status are not reset by an item change.

## Strings (PT-PT, all in `src/taskpane/strings.ts`)
F-001/F-002 strings as approved on 2026-10-22 (D6 = PT-PT), plus: `title` "DIMSCALE Mail"; tabs "Analisar", "Responder"; `comingSoon` "IA (resumo e resposta sugerida) em breve."; `assurance` "Sem envio automático. O envio é sempre feito por si."; `responderSoon` "Respostas automáticas para contactos escolhidos: em breve."; footer labels from the status table; accessible names "Pré-visualização do corpo", "A carregar a mensagem…". European Portuguese wording ("iniciar sessão", "ligar", "e-mail"). `html lang="pt-PT"`. Manifest ribbon strings "DIMSCALE MAIL" are a product name and stay as is.

## Acceptance criteria
"Mocks" = vitest/jsdom with mocked Office.js. "Visual" = a person in a browser. "Live" = Outlook (classic on Windows Version 2409+ unless stated), build recorded.

### Preserved behavior
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-1 | Every F-001 AC (1–10) and every pane-related F-002 AC (1–7, 16, 17, 18, 20) still holds, with the PT-PT literals approved on 2026-10-22 replacing the English ones. No assertion removed or weakened; the vitest count does not drop below the step-0 baseline; `deno test` unchanged. | Mocks |
| F-012/AC-2 | The pane reads only `subject`, `from`, `dateTimeCreated`, `body` from the item (Proxy test kept). The only new Office.js calls are the theme calls (AC-7), `ItemChanged` registration (AC-35) and `requirements.isSetSupported`. No new network request; add-in permission stays `ReadItem`; no change under `supabase/`, `auth.ts`, `backend.ts`, `config.ts`. | Mocks, diff |
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
| F-012/AC-10 (amended) | Production renders exactly the content set: header, tabs, Analisar (subject, sender and date, the "Mensagem" accordion **collapsed on open**, `comingSoon`, assurance), Responder (placeholder card until F-013's UI step, then the F-013 UI), footer, notices when needed. No other section, no disabled fake control. | Mocks |
| F-012/AC-11 (amended) | Header is one row ≤ 40 px at 320 px: inline wordmark SVG with an accessible name, the text "Mail", exactly one `h1` named "DIMSCALE Mail"; wordmark black on light, white on dark. | Mocks, Visual |
| F-012/AC-12 (amended) | Message context shows subject, sender (rules above) and date always. The body excerpt is inside a Fluent `Accordion` item "Mensagem" that is **collapsed by default** (the body text is not in the accessibility tree or tab order until the user opens it; opening it needs a click or Enter/Space, never happens on its own, and an item change collapses it again). When opened it shows the excerpt with "(truncado)" when cut and the length line. All F-001 edge cases hold (empty subject/body, no item, loading, body error with Retry and `role="alert"`, large body, emoji, injection text); the body error notice and "Tentar novamente" stay visible **outside** the collapsed accordion. On body error, subject, sender and date are still shown. Body text ≤ 2,000 characters, whitespace kept. F-001 ACs that check the body text are verified after opening the accordion in the test. | Mocks |
| F-012/AC-13 | Backend status renders the six F-002 states per the table, same state machine and handlers, visible text for every state, footer `role="status"`, error notice `role="alert"`, sign-in popup only from the "Iniciar sessão" click. | Mocks |
| F-012/AC-14 | **Superseded** (future-slot components dropped by the compact content set). | — |
| F-012/AC-15 | The assurance line is present in every Analisar state that shows the message, static, not removable by a prop. | Mocks |
| F-012/AC-16 (amended) | Tabs: `TabList` semantics (`role="tablist"`, `aria-selected`, arrow keys), exactly "Analisar" and "Responder", Analisar selected on open; "Escrever" cannot render. Switching tabs keeps the message state. | Mocks |

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
| F-012/AC-26 (amended) | Manifest diff contains only: the ribbon entry of AC-32 (button label "DIMSCALE MAIL" instead of "Abrir"), `SupportsPinning` (AC-34), the icon files if renamed, and a `<Version>` bump. Permission stays `ReadItem`; no `CustomTab` in the production manifest (until AC-40 is decided and the user approves an amendment); `npm run validate` and `manifest.test.ts` pass; `dist/manifest.xml` has the production URLs. | Diff, command output |
| F-012/AC-27 | No new network host (no webfont, CDN, external image, analytics). F-011 checks still pass (hashed names, no source maps, no secrets, no `localhost` in the production manifest). Bundle size delta recorded against the step-0 budget. | Build output, grep |
| F-012/AC-28 | Dependency change limited to `@fluentui/react-icons` as a direct dependency at the version already in the lockfile; `npm audit` recorded with no new High/Critical. | Diff, command output |
| F-012/AC-29 | Typecheck, lint, tests, `npm run build`, `build:dev`, `validate` pass. After merge, the F-011 workflow deploys. | Command output |
| F-012/AC-30 (amended) | Visual acceptance by the user in classic Outlook (then new Outlook and Mac), light and dark, default/narrow/wide pane, for: message loaded (accordion collapsed and open), loading, no message, body error, the reachable backend states. Screenshots without real email content. | Live (user) |
| F-012/AC-31 | **Superseded** (the prototype gate was met: direction D chosen 2026-10-23). | — |

### Ribbon and pinning (new)
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-32 | The manifest's `MessageReadCommandSurface` has exactly one `<OfficeTab id="TabDefault">` with exactly one `<Group>` labelled "DIMSCALE MAIL" containing one `Button` labelled "DIMSCALE MAIL" (Supertip title "DIMSCALE MAIL") whose action is `ShowTaskpane`. No `<CustomTab>` in `manifest.xml` (the test manifests of AC-37 are separate files). | `manifest.test.ts`, `npm run validate` |
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

## Out of scope
- Any AI feature or backend call (F-003 onward). The F-013 contact UI itself is specified and built under F-013.
- `CustomTab` in the production manifest (only the experiment ACs 37 to 40, then a separate approval), compose mode and "Escrever" (F-010), localization machinery, first-run experience, user theme override, persisted preferences, webfonts, analytics.
- Changing `auth.ts`, `backend.ts`, `config.ts` or Supabase functions.

## Risks
- **CustomTab behaviour unknown outside classic Outlook** (documented as unavailable on Mac, web and new Outlook). Production uses only `TabDefault`; the test manifest (AC-37) is separate with another `<Id>`, so a bad result can't break the production install. Two add-ins with different IDs sideloaded together on one client may both show; the user removes the test one afterwards.
- **Classic Mac UI has Mailbox 1.8 only**: no theme API and no 1.14 features there; the pane falls back to `prefers-color-scheme` (AC-42).
- **Theme detection under-documented for Outlook** (values per theme, whether the event fires in classic): safe fallback; AC-8 live.
- **Stale content in a pinned pane** if a late body read wins: AC-35 test.
- **Manifest change needs the admin**: bundled with the planned centralized deployment.
- **Existing tests assert exact strings and DOM shapes**: the main mechanical risk to AC-1.

## Brandbook rules (from the DIMSCALE Brand Book, 2026-10-22; source files stay in the git-ignored `Modelo_DIMSCALE/`)
- The logo is used only in black or white, never coloured. Black wordmark on light, white on dark.
- Digital wordmark minimum about 20 mm (~76 px); clear space = logo height.
- For e-mail and online platforms the brandbook prescribes Helvetica; the pane uses Segoe UI (Outlook-native, direction D) for text and the wordmark as SVG, so no font is bundled.

## User decisions
- 2026-10-22: UI language **PT-PT** (F-001/F-002 literals amended); icon **green "D" on white** (icon-only deviation from the black/white logo rule; exact green from the prototype); "Competências" = "Resumo da conversa" (later feature); pane easy to open/close → **pinnable task pane** (no left docking exists for web add-ins: https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane).
- 2026-10-23: **direction D**, **320 px compact**, **"DIMSCALE MAIL" ribbon entry**, **Responder = contact selection for automatic drafts (F-013)**.

## Open questions
- **Q-R1 (custom tab): answered 2026-10-23, the user wants an own "DIMSCALE MAIL" tab.** Handled by the experiment ACs 37 to 40. Still open: does new Outlook or Mac reject, ignore or fall back when the manifest declares `CustomTab` (undocumented), and does Outlook accept `CustomTab` with `OfficeTab` in one extension point (undocumented; variant A is the fallback)?
- **Q-R2: answered 2026-10-23, "DIMSCALE MAIL".** Closed.
- **Q-R3.** How does the user sideload a test manifest on Mac (and in new Outlook)? The testing-and-tips page is cited but not checked per client.
- **Q-R4.** Does a pinned pane reopen after an Outlook restart (AC-41)? Undocumented.
- **Q-R5.** Does `OfficeThemeChanged` fire on new Mac UI (AC-42)? Requirement set says supported, firing not documented.
- **Q11 (kept).** Does Outlook show the add-in name above the pane? If yes, the header may keep only the wordmark.
- **Q12 (kept).** Office theme values per theme, `OfficeThemeChanged` in classic, `forced-colors` in the webview: answered by AC-8, AC-9.
- **Q17 (kept).** Real pane widths in classic Outlook: measure live.
