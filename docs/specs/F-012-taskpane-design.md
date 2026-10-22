# F-012: Task pane UI/UX design (DIMSCALE brand, light and dark, shell for the AI sections)

- **Status:** Draft (2026-10-22). Not approved. Several decisions in [Open questions](#open-questions) need the user before the plan can start.
- **Tier:** Active item, presentation only. No Graph, no new Office.js permission, no new backend call, no AI.
- **Depends on:** F-001 (Done), F-002 (Done), F-011 (Done: the pane is served from GitHub Pages, so a UI-only change reaches users without a manifest re-upload).
- **Plan:** [docs/plans/F-012-taskpane-design.md](../plans/F-012-taskpane-design.md)
- **Working agreement:** the user was told before design work started (status.md) and asked for it, wanting top-quality UI/UX.
- **Reference supplied by the user:** a Portuguese mockup (tabs Analisar / Responder / Escrever; sections Resumo inteligente, Detalhes extraídos, Resposta sugerida, Ferramentas rápidas, Atividade IA; orange primary button) and a 32 px "green D on black" icon. The mockup image itself was not available to the author of this spec, only the description given by the orchestrator. Colors and spacing below are therefore proposals to be checked against the image in the prototype step.

## Goal
Replace the generator-default look of the pane with a polished, accessible, DIMSCALE-branded interface that:
1. keeps everything F-001 and F-002 already do (message context, backend status);
2. looks and behaves like a native Outlook companion (Fluent UI React v9, light and dark);
3. provides the layout "slots" where F-003 (summary), F-004/F-006 (reply), F-005 (open reply form), F-007 (classification) and later tools will appear, **without building them**.

## Design principles
1. **Content over chrome, brand with restraint.** Microsoft: "Favor content over chrome... Brand your experience wisely... keep the focus on content and task completion, not brand attention." The header is one compact row. Brand color is used for primary actions, the wordmark and a few accents, not for large fills (except in the optional Direction B header).
2. **The user is in control, and the UI says so.** No action happens without a click. A static assurance line ("Sem envio automático") is always visible. No UI label may promise something the code does not do (for example "Criar rascunho" before drafts exist).
3. **Email content is untrusted.** Subject, sender, body, names and all future AI output are rendered as text nodes. Nothing in the redesign introduces `innerHTML`, links built from email content, or images from email content.
4. **Show only what works.** A section whose feature is not built is not shown in the production pane (decision proposed in Q5). The full layout exists as components and in the prototype.
5. **Accessible by construction.** WCAG 2.1 AA contrast, full keyboard use, screen reader names and roles, no meaning by color alone, reflow at 320 px. Microsoft asks add-ins to follow these and to be tested with Narrator, JAWS or NVDA and the four Windows contrast themes.
6. **Native components, branded tokens.** Fluent UI React v9 components restyled through a custom brand theme, not hand-made widgets ("Don't invent new UI components that contradict established Office UI").
7. **Small and quiet.** Compact 4 px grid, 14 px body text, 12 px minimum, no decorative animation, no external fonts or images (F-001/AC-9 stays true).

## Verified platform facts

| Fact | Source |
|---|---|
| Fluent UI React is the recommended design language for React add-ins. "Favor content over chrome." "Brand your experience wisely." Apply document and UI theming where appropriate. Support mouse, keyboard and touch, and be responsive. | [Design the UI of Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/add-in-design) |
| Task pane best practices: include the add-in name in the title, short names, do not append "add-in" or "for Outlook" or the company name; include a navigational or commanding element such as CommandBar or Pivot at the top; a brand bar at the bottom "unless your add-in is to be used within Outlook". Outlook on the web pane size 320x570. The personality menu is not supported in Outlook. | [Task panes in Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/design/task-pane-add-ins) |
| Branding: use familiar components with branding accents (typography, color); "Don't repeat your task pane name in an immediately adjacent brand bar at the top"; use brand elements sparingly; a brand bar footer and a splash screen are the documented patterns. | [Branding patterns](https://learn.microsoft.com/office/dev/add-ins/design/branding-patterns) |
| First-run: keep it simple, be actionable, allow skipping. Patterns: carousel, value placemat, video. (Not needed for an internal add-in with no setup, see Out of scope.) | [First-run experience patterns](https://learn.microsoft.com/office/dev/add-ins/design/first-run-experience-patterns) |
| Accessibility: everything operable by keyboard (Tab and arrows), logical DOM order, ARIA names/roles/states, do not rely on color, shape or location alone, do not move focus unless the user asks, contrast at least 4.5:1 for normal text and 3:1 for large text, check all four Windows contrast themes (Aquatic, Desert, Dusk, Night sky), text alternatives for meaningful icons, test with Narrator, JAWS, NVDA, run Accessibility Insights in CI. | [Accessibility guidelines](https://learn.microsoft.com/office/dev/add-ins/design/accessibility-guidelines) |
| Ribbon icon guidelines (Monoline): PNG, 32-bit with transparent background; required 100% sizes 16, 20 and 32 px, recommended also 24, 40, 48, 64, 80, 96; flat color, no gradients; "Avoid using text characters in icons"; at most two elements; up to one extra color besides gray. | [Icon guidelines](https://learn.microsoft.com/office/dev/add-ins/design/add-in-icons) |
| Manifest `IconUrl`: for a Mail add-in 64 x 64; `HighResolutionIconUrl`: for mail 128 x 128; shown in File > Manage add-ins (Outlook). The ribbon button uses `bt:Image` sizes 16, 32, 80 (already in `manifest.xml`). Formats PNG and others. | [IconUrl](https://learn.microsoft.com/javascript/api/manifest/iconurl), [HighResolutionIconUrl](https://learn.microsoft.com/javascript/api/manifest/highresolutioniconurl) |
| Icon file names are in the manifest. Changing a file name needs a manifest change; replacing the PNG at the same URL does not (but caches can hold the old image, see Risks). Office keeps a web cache and a Wef cache. | [Clear the Office cache](https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache) (as used in F-011) |
| `Office.context.officeTheme` in Outlook: supported from **Mailbox 1.14**, not in event-based-activation add-ins. Members `bodyBackgroundColor`, `bodyForegroundColor`, `controlBackgroundColor`, `controlForegroundColor` are hex triplets. **`isDarkTheme` and `themeId` are not supported in Outlook**: "To determine the current theme in Outlook, use the `body*Color` and `control*Color` properties instead." Theme changes: `Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, handler)`, Mailbox 1.14, read item permission, compose or read, task panes only. `OfficeThemeChangedEventArgs.officeTheme` "gets the updated Office theme". | [officeTheme](https://learn.microsoft.com/javascript/api/office/office.context#office-office-context-officetheme-member), [OfficeTheme](https://learn.microsoft.com/javascript/api/office/office.officetheme), [Mailbox.addHandlerAsync](https://learn.microsoft.com/javascript/api/outlook/office.mailbox), [Mailbox 1.14](https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-14) |
| Classic Outlook on Windows (Microsoft 365 or retail perpetual) supports Mailbox 1.14 from Version 2404 (Build 17530.15000). NAA (F-002) already needs Version 2409, so every client that can sign in also has 1.14. New Outlook, Outlook on the web: 1.14 listed. Outlook 2021 volume-licensed: up to 1.9 only (no theme API, but also no NAA). Check at runtime with `Office.context.requirements.isSetSupported("Mailbox", "1.14")`; the manifest minimum stays 1.1 / 1.3. | [Outlook requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) |
| `displayReplyFormAsync` (Mailbox 1.9, read item, message read mode) "displays a reply form"; the page does not say it saves a draft or that it sends. Not supported in Outlook on Android or iOS. | [MessageRead](https://learn.microsoft.com/javascript/api/outlook/office.messageread) |
| Manifest localization: `DefaultLocale` plus `<Override Locale="pt-pt" .../>` on `DisplayName`, `Description`, `IconUrl`, `HighResolutionIconUrl`, `SourceLocation` and `bt:` resources. A single override per language. | [Localization](https://learn.microsoft.com/office/dev/add-ins/develop/localization) |
| `Office.context.displayLanguage` returns the Office UI language as an RFC 1766 tag, supported in Outlook. | [Office.Context](https://learn.microsoft.com/javascript/api/office/office.context) |
| Fluent UI React v9 custom brand: `createLightTheme(brand)` and `createDarkTheme(brand)` take `BrandVariants` = keys 10, 20, ..., 160 (16 shades, 10 darkest, 160 lightest); `createHighContrastTheme()` exists. In the installed package (`@fluentui/react-components` 9.74.10, `@fluentui/tokens`), the light theme uses brand[80] for `colorBrandBackground`, brand[70] hover, brand[40] pressed, brand[160] for `colorBrandBackground2`, brand[80]/[70] for brand foreground and links; the dark theme uses brand[70] background, brand[80] hover, brand[100]/[110] foreground, brand[20] for `colorBrandBackground2`; `colorNeutralForegroundOnBrand` is white in the light theme. Radii 4/6/8 px, base font sizes 12/14/16 px. Verified in the installed type definitions and token source, not on the Fluent docs site (the Storybook docs page rendered no text to the fetch tool, see Q19). Components exist in the installed package: Accordion, Badge, Card, Divider, MessageBar, Skeleton, Spinner, Tabs/TabList, Tags, Tooltip, Button, Link, Toolbar. | Installed package files, [Fluent UI React](https://react.fluentui.dev), [Fluent UI in Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/quickstarts/fluent-react-quickstart) |
| The Office Fluent quickstart notes the generated `taskpane.html` "tests whether the webview control supports Fluent UI React v9". Fluent v9 needs a modern webview (decision D8 and the IE11 polyfills stay out of this feature). | same quickstart |

## Brand inputs (from the website project, read-only)
- Tokens (`dimscale.newsite/src/app/globals.css`): navy `#102d3f`, navy-deep `#081b24`, blue `#173e54`, ink `#231f20`, muted `#617484`, paper `#fbfcfc`, mist `#f1f4f5`, line `#d9e0e3`, teal `#7fd1c7`. The site also uses `#405969` for secondary text and `#245064` for icons. Fonts: "Helvetica Neue", Helvetica, Arial (the site comment says this is the web fallback for the brandbook font "Rational" "until licensed files exist").
- Mark: the site favicon (`src/app/icon.svg`) is a navy `#102d3f` square with a white outlined "D" (stroke 3 on a 32 px grid). Wordmark: `public/brand/dimscale-wordmark.svg` (ink `#231f20`, 509 x 43, letter-spaced capitals). The brandbook PDF could not be rendered, so clear-space, minimum-size and inverse-logo rules are unknown (Q15).
- Orange is not a brand color. The reference's orange is an alternative the user must decide on (Q2).

## Directions to choose from (Q1)
All three share the same sections, components, theme module and strings. They differ in header, section container and density.

| | Direction A: "Referência" (primary, from the user's mockup) | Direction B: "Navy header" | Direction C: "Cards" |
|---|---|---|---|
| Header | Compact row on the page background: wordmark + "Mail AI", thin divider below | 44 px navy band, white wordmark, teal accents (teal on navy is 8.07:1) | Same as A |
| Page background | White (light) / navy-deep (dark) | Paper / navy-deep | Mist `#f1f4f5` / navy-deep, sections on white or navy Cards |
| Sections | Flat blocks separated by thin dividers, icon on the left of each heading | Same as A | Fluent Card per section, 6 px radius, subtle border |
| Density | Most compact | Compact | Slightly airier (more padding) |
| Brand strength | Medium-low (accent on buttons and active tab) | Highest (navy band) | Medium-low |
| Risk | Plain if the accent is navy | Band can look heavy at 320 px and repeats the pane title Outlook may already show (Q11) | Cards add vertical space, and muted text fails on mist (see contrast) |

Accent variants for A (Q2): **A-navy** (primary button navy `#102d3f`, white text 14.30:1; teal only as a non-text accent on dark surfaces) is brand-faithful and recommended. **A-orange** follows the mockup (see the contrast table: the mockup-like orange fails with white text, so the shippable orange is a darker one).

The recommendation is Direction A with the A-navy accent, taking from B only an optional thin teal line under the header in dark mode. The prototype (plan step 1) renders A (both accents), B and C so the user can pick.

## Design tokens: brand to Fluent theme

### Light theme (`createLightTheme(lightBrand)`)
| Fluent token | Value | Source |
|---|---|---|
| Brand ramp 80 (`colorBrandBackground`, `colorBrandForeground1`, `colorBrandStroke1`) | `#102d3f` navy | brand |
| ramp 70 (hover, `colorBrandForegroundLink`) | darker navy, about `#0c2433` | derived |
| ramp 60 | `#081b24` navy-deep | brand |
| ramp 40 (pressed) | about `#050f15` | derived |
| ramp 90 | `#173e54` blue | brand |
| ramp 160 (`colorBrandBackground2`, tinted surfaces) | `#f1f4f5` mist | brand |
| other steps | interpolated between the anchors, fixed and recorded in plan step 3 | derived |
| `colorNeutralBackground1` | `#ffffff` (page and cards). `#fbfcfc` paper is allowed for the page in Direction B | brand |
| `colorNeutralForeground1` | `#231f20` ink | brand |
| `colorNeutralForeground3` (secondary text) | `#405969` on any surface; `#617484` muted only on white or paper | brand-adjacent, see contrast |
| `colorNeutralStroke2` (dividers) | `#d9e0e3` line | brand |
| Control borders that must be perceivable (inputs, outlined button, reply box) | `#617484` (4.84:1 on white) | brand |
| Accent (non-text only) | teal `#7fd1c7` is **not used on light surfaces** (1.77:1) | see contrast |

### Dark theme (`createDarkTheme(darkBrand)`)
The dark theme needs a different ramp. Fluent's dark theme paints primary buttons with ramp 70; with a navy ramp the button would be navy on navy. So the dark ramp is built around teal.
| Fluent token | Value |
|---|---|
| Ramp 70 (`colorBrandBackground`, primary button) | `#7fd1c7` teal |
| Ramp 80 (hover) | lighter teal, about `#96dad1` |
| Ramp 40 (pressed) | about `#5fb9ae` |
| Ramp 100 and 110 (links, brand foreground) | about `#a6e0d9` and `#bbe8e2` (11.98:1 on navy-deep for `#a6e0d9`) |
| Ramp 20 (`colorBrandBackground2`) | dark teal tint, about `#0d3a3f` |
| Override `colorNeutralForegroundOnBrand` | `#081b24` (white on teal fails, navy-deep on teal is 9.93:1) |
| `colorNeutralBackground1` (page) | `#081b24` navy-deep |
| Surfaces (cards, reply box, message preview) | `#102d3f` navy |
| `colorNeutralForeground1` | `#fbfcfc` paper |
| Secondary text | `#9fb2bf` (derived, not a brand token): 8.05:1 on navy-deep, 6.53:1 on navy. The brand muted `#617484` is **not used on dark** (3.64:1 on navy-deep) |
| Dividers | `#173e54` blue (decorative) |

### High contrast
When `matchMedia("(forced-colors: active)")` matches, use Fluent's `createHighContrastTheme()` and drop custom colors. Whether the Outlook webview honors forced colors is not documented (Q12, AC-9 checks it live).

### Typography
- Body and controls: Segoe UI Variable / Segoe UI (native to Windows Outlook) falling back to "Helvetica Neue", Helvetica, Arial, sans-serif. Section headings and the title: the brand stack "Helvetica Neue", Helvetica, Arial (renders as Arial on Windows, since Helvetica Neue is not installed there). The Rational font has no licensed web files, and no webfont may be loaded (F-001/AC-9). Q10 asks whether to use the brand stack everywhere.
- Sizes: title 16 px semibold, section heading 14 px semibold, body 14 px, meta and chips 12 px (the floor). Line height from Fluent tokens.

### Spacing, shape, motion
4 px grid; pane padding 12 px at 320 and 16 px from 400; section vertical gap 12 px; corner radius 6 px (`borderRadiusLarge`) for buttons, chips, inputs and cards ("rounded buttons" of the reference); one shadow level at most (Direction C only). Motion: Fluent defaults, no custom animation, and none at all under `prefers-reduced-motion: reduce`.

### Icons in the UI
`@fluentui/react-icons` (Fluent System Icons, 20 px regular for section headings, 16 px for inline). It is already installed transitively (2.0.344, dependency of the Fluent component packages); F-001 removed the direct dependency because it was unused. F-012 re-adds it as a direct dependency, with named per-icon imports so the bundle only contains the icons used. Candidate icons (names confirmed to exist as files in the installed package; exact export names are confirmed by the coder): sparkle (Resumo inteligente), document-bullet-list (Detalhes extraídos), chat-multiple or comment (Resposta sugerida), wrench (Ferramentas rápidas), shield-checkmark (Sem envio automático), arrow-clockwise (Tentar novamente/Regenerar), mail (Mensagem), person-key (Iniciar sessão), chevron-down, checkmark-circle, warning, error-circle, translate, text-grammar-wand, history, calendar-clock.

## Contrast table (WCAG 2.1 AA)
Ratios are computed with the WCAG relative-luminance formula on 2026-10-22, by hand, so they are good to about two decimals. Plan step 3 adds a test that recomputes every row from the theme tokens and fails below the threshold (AC-6). Thresholds: 4.5 for normal text, 3 for large text (18.66 px bold or 24 px) and for non-text UI parts (WCAG 1.4.11).

Relative luminance used: white 1.000, paper 0.972, mist 0.900, line 0.736, teal 0.542, muted 0.167, `#405969` 0.093, `#9fb2bf` 0.430, blue 0.043, navy 0.0234, ink 0.0144, navy-deep 0.0096, `#F26B21` 0.295, `#C2410C` 0.153.

| # | Foreground on background | Use | Ratio | Result |
|---|---|---|---|---|
| 1 | ink `#231f20` on white | body text, light | 16.30 | Pass (AAA) |
| 2 | ink on paper `#fbfcfc` | body text on paper | 15.86 | Pass |
| 3 | navy `#102d3f` on white | headings, outline button text | 14.30 | Pass |
| 4 | white on navy `#102d3f` | primary button label (A-navy) | 14.30 | Pass |
| 5 | white on blue `#173e54` | primary button hover | 11.33 | Pass |
| 6 | muted `#617484` on white | secondary text | 4.84 | Pass (AA only) |
| 7 | muted on paper | secondary text on paper | 4.71 | Pass |
| 8 | muted on mist `#f1f4f5` | secondary text on Direction C background | **4.38** | **Fail**: not allowed on mist |
| 9 | `#405969` on mist | secondary text on mist | 6.66 | Pass (use this on mist) |
| 10 | `#405969` on white | secondary text, safe default | 7.37 | Pass (AAA) |
| 11 | muted `#617484` as a border on white | control borders | 4.84 | Pass (>= 3) |
| 12 | line `#d9e0e3` on white | dividers | 1.34 | Decorative only. Never the only boundary of a control |
| 13 | teal `#7fd1c7` on white | text or UI part, light | **1.77** | **Fail** (text and non-text). Not used on light surfaces |
| 14 | teal on paper / mist | same | 1.72 / 1.60 | **Fail** |
| 15 | teal on navy `#102d3f` | accents on navy (Direction B header) | 8.07 | Pass |
| 16 | teal on navy-deep `#081b24` | accents and links, dark | 9.93 | Pass |
| 17 | navy-deep on teal | label on a teal button (dark primary) | 9.93 | Pass |
| 18 | paper `#fbfcfc` on navy-deep | body text, dark | 17.13 | Pass |
| 19 | paper on navy `#102d3f` | text on dark surfaces | 13.91 | Pass |
| 20 | `#9fb2bf` on navy-deep / navy | secondary text, dark | 8.05 / 6.53 | Pass |
| 21 | muted `#617484` on navy-deep | secondary text, dark | **3.64** | **Fail**: not used on dark |
| 22 | `#a6e0d9` on navy-deep | links, dark | 11.98 | Pass |
| 23 | white on mockup-like orange `#F26B21` | primary button (A-orange as in the mockup) | **3.04** | **Fail** for 14 px labels (passes only as large text) |
| 24 | navy `#102d3f` on `#F26B21` | primary button with dark label | 4.70 | Pass (barely) |
| 25 | white on darker orange `#C2410C` | primary button (A-orange, shippable) | 5.18 | Pass |
| 26 | `#C2410C` on white | orange text or active-tab underline, light | 5.18 | Pass |
| 27 | `#F26B21` on navy-deep | orange accent, dark | 5.79 | Pass |
| 28 | `#F26B21` on white as an underline or border | non-text | 3.04 | Pass at 3:1 (marginal) |

The real orange of the mockup is unknown (only the word "orange" reached the author), so rows 23 to 28 are for a typical orange; the test re-checks the final value.

Status, chip and message-bar colors (info, success, warning, error, "priority") use Fluent's semantic palette tokens; the prototype step measures every pair actually rendered, and the AC-6 test covers them. No state is shown by color alone (AC-21).

## Layout
Everything below is for a pane 320 to 500 px wide, one column, vertical scroll only (no horizontal scroll at 320 px). Above 500 px the content column is capped at 640 px and left-aligned.

### Slots (the structure that hosts the AI features)
Order top to bottom. "Prod" is what F-012 shows in the production pane (Q5 recommendation).

| Slot id | PT-PT heading | Icon | Filled by | Prod in F-012 | Notes for the future spec |
|---|---|---|---|---|---|
| `header` | DIMSCALE Mail AI | wordmark | F-012 | Shown | One `h1`. |
| `notices` | none | per intent | F-002 status (sign-in needed, error, unsupported, not configured) and later AI errors | Shown when needed | Fluent MessageBar. |
| `tabs` | Analisar, Responder (Escrever) | none | F-012 component | Hidden (only one view exists) | Shown when two or more views are available (Q6). `Escrever` is compose-mode only (F-010) and is never rendered in read mode. |
| `message` | Mensagem | mail | F-001 | Shown | Collapsible, expanded by default, state in memory only. |
| `summary` | Resumo inteligente | sparkle | F-003 (3 bullets), F-007 (tag chips) | Hidden | Plain text bullets. Chips are suggestions: show "Sugestões da IA, podem estar erradas" (product rule: classification is never certain). Show the external-AI notice from security.md once (one line). |
| `details` | Detalhes extraídos | document list | **new feature, not in the backlog** (proposed F-014, needs a requirements.md entry) | Hidden | Label/value rows (Prazo, Projeto, Ação necessária). Extra content sent to the AI, so it needs its own minimization limit. |
| `reply` | Resposta sugerida | chat | F-004 (text), F-006 (Ajustar tom, Regenerar), F-005 (open the reply form) | Hidden | Bordered read-only text box, buttons below. Primary button label is "Abrir resposta para revisão" (matches F-005: `displayReplyFormAsync` opens a form and the user sends). "Criar rascunho no Outlook" is only correct with F-009 (Graph, `Mail.ReadWrite`, which the admin was just asked to revoke) or if a live test proves the form is saved as a draft (the docs do not say). |
| `tools` | Ferramentas rápidas | wrench | each tool has its own feature: Resumir (F-003), Extrair dados (details), Traduzir, Verificar tom, Seguimento, Competências, Melhorar texto (compose, F-010) | Hidden | Tiles appear one by one as their feature ships. A tile with no feature is never shown disabled. "Competências" is unclear (Q7). "Seguimento" must stay a text suggestion: reminders or flags would need Graph write access. |
| `assurance` | Sem envio automático | shield-checkmark | F-012, static | Shown | One line: "Sem envio automático. O envio é sempre feito por si." The reference's "Atividade IA" log is not built (a log of events needs a design for what is stored, and nothing may be stored now). |
| `footer` | none | dot | F-002 | Shown | Status dot and text (see states), `role="status"`. |

When no AI slot is shown, the pane shows a single compact info card (`comingSoon`): "As ações de IA (resumo e resposta sugerida) chegam em breve." Plain text, no disabled buttons (Q5).

### Wireframe 1: production pane, message loaded, Direction A, light (320 px)
```
+--------------------------------------+
| DIMSCALE  Mail AI                    |  header 44 px, h1
+--------------------------------------+
| (notices: empty)                     |
| [v] (mail) Mensagem                  |  accordion header, button
|   Assunto                            |
|   Pedido de orcamento - Obra 203     |  wraps, break-word
|   De    Ana Silva                    |
|         <ana@exemplo.pt>             |
|   Data  02/01/2026 03:04             |
|   Pre-visualizacao                   |
|   +--------------------------------+ |
|   | Bom dia,                       | |  pre-wrap, max-height ~160 px,
|   | Segue o pedido de...           | |  focusable scroll region
|   +--------------------------------+ |
|   (truncado)  Corpo: 2500 caracteres |
+--------------------------------------+
| (sparkle) As acoes de IA chegam em   |  comingSoon card
| breve.                               |
+--------------------------------------+
| (shield) Sem envio automatico.       |  assurance
|          O envio e sempre feito por  |
|          si.                         |
+--------------------------------------+
| (o) Backend ligado como Ana Silva    |  footer, 12 px
+--------------------------------------+
```

### Wireframe 2: full layout (prototype and the future pane), Direction A, light (320 px)
```
+--------------------------------------+
| DIMSCALE  Mail AI                    |
+--------------------------------------+
| Analisar | Responder                 |  TabList (Escrever absent in read mode)
|=========                             |  navy underline (selected)
| [v] (mail) Mensagem                  |
|   Pedido de orcamento - Obra 203     |
|   Ana Silva - 02/01/2026 03:04       |
+--------------------------------------+
| (sparkle) Resumo inteligente         |
|  - Pedido de orcamento para obra     |
|  - Prazo ate dia 15                  |
|  - Pede visita tecnica               |
|  [Pedido de orcamento] [Prioridade   |
|   alta] [Tom neutro]                 |
|  Sugestoes da IA, podem estar erradas|
+--------------------------------------+
| (list) Detalhes extraidos            |
|  Prazo            15 de novembro     |
|  Projeto          Obra 203           |
|  Acao necessaria  Enviar proposta    |
+--------------------------------------+
| (chat) Resposta sugerida             |
|  +--------------------------------+  |
|  | Bom dia Ana, obrigado pelo     |  |
|  | contacto. ...                  |  |
|  +--------------------------------+  |
|  [ Abrir resposta para revisao  ]    |  primary, full width at <400 px
|  [Regenerar] [Ajustar tom]           |  secondary (outline)
+--------------------------------------+
| (wrench) Ferramentas rapidas         |
|  [Resumir] [Traduzir] [Extrair dados]|  tiles, 3 columns <400 px,
|  [Verificar tom] [Seguimento] ...    |  4 columns from 400 px
+--------------------------------------+
| (shield) Sem envio automatico.       |
+--------------------------------------+
| (o) Backend ligado como Ana Silva    |
+--------------------------------------+
```
In the production F-012 pane only the header, message, assurance, comingSoon card and footer exist (Wireframe 1). Sample text above is fake and exists only in the prototype.

### Wireframe 3: loading
```
+--------------------------------------+
| DIMSCALE  Mail AI                    |
+--------------------------------------+
| [v] (mail) Mensagem                  |
|   ############                       |  Skeleton lines (subject, sender,
|   ########  ##########               |  date), aria-busy="true",
|   +--------------------------------+ |  no spinner text spam
|   | ############################## | |
|   | ##########################     | |
|   +--------------------------------+ |
+--------------------------------------+
| (o) A ligar ao backend DIMSCALE...   |
+--------------------------------------+
```
Accessible name: the region carries `aria-busy` and a visually hidden "A carregar a mensagem..." (F-001 loading text).

### Wireframe 4: no message selected
```
+--------------------------------------+
| DIMSCALE  Mail AI                    |
+--------------------------------------+
|                                      |
|        (mail icon, 32 px)            |
|   Abra uma mensagem para usar o      |
|   DIMSCALE Mail AI                   |
|                                      |
+--------------------------------------+
| (o) Backend ligado como Ana Silva    |
+--------------------------------------+
```

### Wireframe 5: body could not be read
```
+--------------------------------------+
| DIMSCALE  Mail AI                    |
+--------------------------------------+
| [x] Nao foi possivel ler o corpo da  |  MessageBar intent=error, role=alert
|     mensagem. Tente novamente.       |
|              [Tentar novamente]      |
+--------------------------------------+
| Assunto, De, Data are still shown if |
| the item has them (they do not need  |
| the body call)                       |
```
Today the F-001 panel shows only the error and Retry. Showing the header fields that exist without the body is allowed because `subject`, `from` and `dateTimeCreated` are already read synchronously. This is a small behavior improvement and is covered by AC-12.

### Backend status states (footer, plus `notices` for the actionable ones)
Status text is always visible text next to a dot or icon. F-002 logic (`auth.ts`, `backend.ts`, `config.ts`, the state machine in `BackendStatus.tsx`) does not change; only its presentation does.

| State | Footer (dot + text) | Notice above the message |
|---|---|---|
| connecting | grey pulse-free dot, "A ligar ao backend DIMSCALE..." | none |
| connected | green check icon, "Backend ligado como {name}" (or "Backend ligado") | none |
| needSignIn | amber icon, "Sessão não iniciada" | MessageBar info: "Inicie sessão para ligar ao backend DIMSCALE" + button "Iniciar sessão" (calls the existing `signIn`, only on click) |
| error | red icon, "Erro de ligação" + code in parentheses when there is one | MessageBar error: the F-002 message (sign-in failed, rejected, unreachable) + code + button "Tentar novamente" |
| unsupported | amber icon, "Sessão indisponível" | MessageBar info: the F-002 unsupported text, no button |
| notConfigured | amber icon, "Backend não configurado" | MessageBar info: the F-002 not-configured text, no button |

(The wireframes drop accents only to stay plain ASCII. The status wording in this table is the PT-PT proposal; with D6 = EN the footer texts are "Connecting...", "Backend connected", "Not signed in", "Connection error", "Sign-in unavailable", "Backend not configured".)

The error code (MSAL `errorCode` or HTTP status) is kept because F-002/AC-4, 5, 7 require it. It is not an email field, so showing it is safe.

### Responsive rules
- 320 px: one column, 12 px padding, tiles in 3 columns, buttons stacked full width, header fields stacked (label above value). 400 px and up: 16 px padding, tiles in 4 columns, label and value on one line, primary and secondary buttons side by side if they fit. The breakpoint (about 400 px) is fixed in the prototype.
- Every text element that holds email content wraps (`overflow-wrap: anywhere`) and has `dir="auto"` so right-to-left mail displays correctly.
- Body preview: `white-space: pre-wrap`, max height about 160 px with its own scroll, `tabindex="0"`, `role="region"`, accessible name "Pré-visualização do corpo". The text limit stays at the F-001 value (2,000 characters).
- Nothing relies on hover (touch and keyboard). Tooltips only repeat visible information.
- Targets at least 32 px high for buttons and tiles (WCAG 2.2 asks 24 px; the Fluent default is 32 px).

### Dark and light theme resolution
1. If `Office.context.requirements.isSetSupported("Mailbox", "1.14")` and `Office.context.officeTheme.bodyBackgroundColor` is a valid hex triplet, theme is dark when its relative luminance is below 0.179 (the point where white and black text have equal contrast), else light. `isDarkTheme` and `themeId` are never used (not supported in Outlook).
2. Otherwise `matchMedia("(prefers-color-scheme: dark)")`.
3. Otherwise light.
4. Re-evaluate on `OfficeThemeChanged` (registered with `Office.context.mailbox.addHandlerAsync` when 1.14 is supported, and removed with `removeHandlerAsync` on unmount) and on the `matchMedia` change event.
5. `forced-colors: active` overrides all of the above with the high-contrast theme.
The pane fills its own background from the theme (it does not inherit Outlook's). Whether Outlook's own dark setting for the reading pane and the Office theme agree, and whether `OfficeThemeChanged` fires in classic Outlook on Windows, are not documented and are open (Q12, AC-8). Before React renders, only the `body` background follows `prefers-color-scheme` so there is no white flash. The `#container` stays empty until `Office.onReady` (F-001 edge case and its test).

## Add-in icons
- Files: `assets/icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png`, same names and paths, so `manifest.xml` and `dist/manifest.xml` need no change. Today they are the generator's gray placeholder circle with the word "Logo". PNG, 32-bit, exact pixel sizes. 64 and 128 are the manifest `IconUrl` and `HighResolutionIconUrl` for a mail add-in; 16, 32 and 80 are the ribbon `bt:Image` entries. The 20, 24, 40, 48 and 96 px sizes in Microsoft's list are optional extras and are not in the manifest, so they are out of scope.
- Design: the DIMSCALE site mark, a navy `#102d3f` full-bleed square with the white outlined "D" (the favicon geometry). It is opaque, so it reads on light and dark ribbons. A small teal `#7fd1c7` sparkle at the top right is an option for the 64, 80 and 128 px sizes (teal on navy is 8.07:1); not at 16 and 32 px where it would be noise. Microsoft's ribbon guidance says to avoid text characters and use a monoline outline style; the "D" is a logo mark, not text, and a filled tile is chosen on purpose for the dark ribbon. This is a documented deviation.
- Source of truth: an SVG master committed in the repo, with the PNGs generated from it (no new dependency; headless Edge or any rasterizer already on the machine, hand-tuned at 16 px if needed).
- The 32 px "green D on black" the user sent is off-brand (green and black are not brand colors). It is treated as a placeholder unless the user says otherwise (Q3).

## Manifest review
- `DisplayName` "DIMSCALE Mail AI": keep. It contains the product name and follows Microsoft's naming advice (no "add-in", no "for Outlook").
- `Description` "Helps you read, triage and answer email with AI.": keep for now.
- Supertip "Shows the context of the open message.": truthful today, keep. It must change only when F-003 ships (never promise AI before it exists).
- Proposed texts for later, to be bundled in one manifest update with the first AI feature, with a `<Version>` bump and an admin re-upload (F-011 runbook): Description "Ajuda a ler, triar e responder a e-mails com IA." (if PT-PT); button label "Abrir DIMSCALE Mail AI"; tooltip "Resumo e resposta sugerida para a mensagem aberta." If the user wants Portuguese in the ribbon, the manifest needs `<Override Locale="pt-pt" .../>` entries, which is also a manifest change.
- Therefore **F-012 changes no field in `manifest.xml`** (AC-26). Icons are replaced at the same URLs. A manifest change would force an admin re-upload ("Manifest changed later" in F-011).

## Strings (only if D6 = PT-PT, Q4)
All strings stay in `src/taskpane/strings.ts`. European Portuguese wording (for example "iniciar sessão", "ligar", "e-mail", "rascunho"), not Brazilian ("entrar", "conectar"). `html lang` becomes `pt-PT`.

| Key | EN today | PT-PT proposal |
|---|---|---|
| title | DIMSCALE Mail AI | DIMSCALE Mail AI |
| noItem | Open a message to use DIMSCALE Mail AI | Abra uma mensagem para usar o DIMSCALE Mail AI |
| loading | Loading message... | A carregar a mensagem… |
| noSubject | (no subject) | (sem assunto) |
| noBody | This message has no text content | Esta mensagem não tem conteúdo de texto |
| truncated | (truncated) | (truncado) |
| error | Could not read the message body. Please try again. | Não foi possível ler o corpo da mensagem. Tente novamente. |
| retry | Retry | Tentar novamente |
| subject / from / date / body | Subject / From / Date / Body | Assunto / De / Data / Pré-visualização |
| bodyLength(n) | Body length: n characters | Comprimento do corpo: n caracteres |
| backendWorking | Connecting to DIMSCALE backend… | A ligar ao backend DIMSCALE… |
| connected / connectedAs | Backend connected (as {name}) | Backend ligado (como {name}) |
| unsupported | This Outlook version doesn't support sign-in for DIMSCALE Mail AI. Update Outlook or use Outlook on the web. | Esta versão do Outlook não suporta o início de sessão do DIMSCALE Mail AI. Atualize o Outlook ou utilize o Outlook na Web. |
| needSignIn / signIn | Sign in to connect to the DIMSCALE backend / Sign in | Inicie sessão para ligar ao backend DIMSCALE / Iniciar sessão |
| signInFailed / signInError | Sign-in failed / Could not get a sign-in token | Falha ao iniciar sessão / Não foi possível obter o token de início de sessão |
| rejected / unreachable | Backend rejected the sign-in / Backend unreachable | O backend rejeitou o início de sessão / Backend inacessível |
| notConfigured | DIMSCALE backend sign-in is not configured yet | O início de sessão no backend DIMSCALE ainda não está configurado |
| new: footer short labels | none | A ligar ao backend DIMSCALE… / Sessão não iniciada / Erro de ligação / Sessão indisponível / Backend não configurado (the connected text is the existing one) |
| new: accessible names | none | Pré-visualização do corpo (preview region); Mensagem (section); "A carregar a mensagem…" (visually hidden, loading) |
| new: assurance | none | Sem envio automático. O envio é sempre feito por si. |
| new: comingSoon | none | As ações de IA (resumo e resposta sugerida) chegam em breve. |
| new: section headings (prototype and future) | none | Resumo inteligente, Detalhes extraídos, Resposta sugerida, Ferramentas rápidas |
| new: reply button (future, F-005) | none | Abrir resposta para revisão |
| new: activity line after F-005 (future) | none | Resposta aberta para revisão (not "Resposta criada em Rascunhos", see the slot table) |

If D6 stays EN, the English column is kept and the Portuguese column is the translation file to add later.

## Acceptance criteria
"Mocks" = verifiable offline (vitest with jsdom and mocked Office.js, grep, file checks, build output). "Visual" = looked at by a person in the prototype or the built pane in a normal browser. "Live" = needs classic Outlook on Windows (Version 2409 or later, as for F-002), with the build recorded.

### Preserved behavior
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-1 | Every F-001 AC (1 to 10) and every F-002 AC that concerns the pane (1 to 7, 16, 17, 18, 20) is still true. Existing tests are kept; only selectors, expected strings (if D6 = PT-PT, per the table) and DOM structure may be edited, and no assertion is removed or weakened. The number of vitest tests does not drop below the baseline recorded in plan step 0 (59 at F-011), and `deno test` (24) is unchanged. Under D6 = PT-PT the English literals in F-001/AC-5, 6, 7 and F-002/AC-1, 7 are replaced by the table values: this is a spec amendment the user approves with this spec. | Mocks (command output, test diff review) |
| F-012/AC-2 | The pane reads only `subject`, `from`, `dateTimeCreated` and `body` from the item (the existing Proxy test is kept and still passes), makes no new Office.js call except the theme calls in AC-7, no new network request, and adds no permission. `git diff` shows no change under `supabase/`, `src/taskpane/auth.ts`, `backend.ts`, `config.ts`. | Mocks (test, diff) |
| F-012/AC-3 | All email-derived and name text (subject, sender name and address, date, body, signed-in name) and all placeholder or future AI text are React text nodes. No `dangerouslySetInnerHTML`, `innerHTML`, `insertAdjacentHTML` or `href` built from email content in `src/`. The F-001 injection tests pass against the new components. | Mocks (tests, grep) |
| F-012/AC-4 | No `console.*` in new code, and no `localStorage`, `sessionStorage`, `document.cookie` or IndexedDB use (the origin is shared org-wide, see F-011). Collapse state and theme live in memory only. | Mocks (grep) |

### Brand, theme, contrast
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-5 | A theme module exports `lightTheme` and `darkTheme` built with `createLightTheme` and `createDarkTheme` from the DIMSCALE ramps of this spec. A test asserts: light `colorBrandBackground` = `#102d3f`; dark `colorBrandBackground` = `#7fd1c7`; dark `colorNeutralForegroundOnBrand` = `#081b24`; page backgrounds white and `#081b24`. Direction and accent decided in Q1 and Q2 are applied. | Mocks |
| F-012/AC-6 | A test computes the WCAG ratio of every pair in the contrast table that the final design uses, reading the colors from the theme tokens, and fails below 4.5 (text) or 3 (non-text, large text). It also asserts that teal is not used as text or as a boundary on any light surface, and that muted `#617484` is not used on mist or on dark surfaces. The semantic status, chip and message-bar pairs rendered in the prototype are measured and listed in the plan evidence. | Mocks |
| F-012/AC-7 | `resolveTheme`/the theme hook follows "Dark and light theme resolution": tests for (a) Mailbox 1.14 supported and a dark `bodyBackgroundColor` gives dark; (b) light gives light; (c) 1.14 unsupported and `prefers-color-scheme: dark` gives dark; (d) neither gives light; (e) a missing or invalid hex falls back to step 2; (f) `forced-colors: active` gives the high-contrast theme; (g) the `OfficeThemeChanged` handler is registered through `addHandlerAsync`, changes the theme when it fires, and is removed on unmount; (h) a mock with no `Office.context.officeTheme` and no `requirements.isSetSupported` does not throw (the existing `index.test.tsx` mock must keep working); (i) `isDarkTheme` and `themeId` do not appear in `src/`. | Mocks |
| F-012/AC-8 | Live, classic Outlook on Windows: for each Office theme the client offers (record the list: expected Colorful, Dark Gray, Black, White, and a system setting if present), the pane is legible, uses the right theme, and has no seam against Outlook that the user considers a defect. Record the `bodyBackgroundColor` value seen for each theme and whether changing the theme with the pane open updates it or needs the pane reopened. | Live (user), recorded in the plan |
| F-012/AC-9 | Live: with each of the four Windows contrast themes (Aquatic, Desert, Dusk, Night sky) all text, buttons, focus indicators and the status icons are visible. Record whether the webview applies `forced-colors`. | Live (user) |

### Layout and components
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-10 | A layout component renders the slots in the order of the slot table. A slot that is disabled renders nothing: no heading, no empty container, nothing in the accessibility tree. Production enables `header`, `notices` (when needed), `message`, `comingSoon`, `assurance`, `footer`; the others stay disabled by one constant. A test renders with every slot on and with the production set. | Mocks |
| F-012/AC-11 | The header is one row at most 48 px high at 320 px, with the DIMSCALE wordmark (as an inline SVG with an accessible name) and the text "Mail AI", exactly one `h1` in the page whose accessible name reads "DIMSCALE Mail AI", no tagline, no second brand element, and no duplicate of the title in the footer. The wordmark follows the brand colors (ink on light, paper on dark). | Mocks (one `h1`, name) and Visual |
| F-012/AC-12 | The message section shows subject, sender (name and address, or the address alone when there is no name, never `<>` or "undefined"), date labelled "Data", body preview with the "(truncado)" marker when cut, and the length line, with all F-001 edge cases: empty or whitespace subject and body, no item, loading (Skeleton, `aria-busy`), body error with Retry (`role="alert"`, no logging), large body, emoji and non-ASCII, injection text. If the body call fails, subject, sender and date that the item has are still shown above the error. The body element still keeps whitespace and holds at most 2,000 characters of text. | Mocks |
| F-012/AC-13 | The backend status renders the six F-002 states as in the status table, with the same state machine and handlers, visible text for every state (never color or icon alone), `role="status"` for the footer, `role="alert"` for the error notice, the signed-in name as a text node, the sign-in popup still opened only from the "Iniciar sessão" click, and Retry calling the same `run`. The F-002 tests for AC-1, 4, 5, 7 pass with only selector or string edits. | Mocks |
| F-012/AC-14 | Presentational components exist for the future slots (`summary`, `details`, `reply`, `tools`, `assurance`) with a props-only API (no Office.js, no network, no state of their own except UI toggles), and `empty`, `loading`, `error` and `ready` variants, each rendering its content as text. They render in the prototype and in component tests, and are not wired to any data in F-012. A button handler prop that is not given means the button is not rendered. No tile or button exists for a feature that has no spec. The `summary` chips always show the "may be wrong" line when a classification chip is present. | Mocks |
| F-012/AC-15 | The assurance line ("Sem envio automático...") is present in every production state that shows the message section, is static text (no logic, no data), and is not removable by a prop. | Mocks |
| F-012/AC-16 | The tab component renders only the views passed to it, uses `TabList` semantics (`role="tablist"`, `aria-selected`, arrow-key movement), is not rendered when fewer than two views exist, and the `Escrever` view cannot be rendered in read mode (no such view is registered). | Mocks |

### Responsive and accessibility
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-17 | At widths 320, 360, 450 and 500 px, with a 200-character unbroken subject, a 120-character address and a 2,000-character body with long lines, there is no horizontal scroll and nothing is clipped or overlapping. Tiles are 3 columns below the breakpoint and 4 above it; the primary button is full width below it. jsdom cannot test layout, so this is checked in a real browser. | Visual (prototype and built pane at fixed widths), then Live at the default pane width and a narrow and a wide pane |
| F-012/AC-18 | Reflow and text size: content reflows at 320 CSS px with no loss (WCAG 1.4.10) and stays usable with text enlarged to 200% (WCAG 1.4.4). | Visual (browser zoom), Live |
| F-012/AC-19 | Keyboard: every interactive element is reachable with Tab in the visual order, tabs move with arrow keys, Enter and Space activate buttons, the body preview scroll area is focusable and scrolls with the keyboard, focus is visible (at least 3:1 against its background) and never trapped or moved unless the user acts. | Mocks (user-event tab order and key tests), Live (manual) |
| F-012/AC-20 | Semantics: landmarks (`header`, `main`, `footer`), one `h1`, an `h2` per section with `aria-labelledby` on the section, decorative icons `aria-hidden`, informative icons with a name, the wordmark with a name, `lang` set on `html`, loading with `aria-busy`, statuses and errors through `role="status"` and `role="alert"`. Read with Windows Narrator in classic Outlook (and NVDA or JAWS if the user has them): record what is read for each state. | Mocks (`getByRole` tests), Live (manual, recorded) |
| F-012/AC-21 | No state is conveyed by color alone: each status, chip and message bar has text, and the status icons differ in shape. | Mocks (DOM checks) and Visual (grayscale screenshot) |
| F-012/AC-22 | Motion: no custom animation; under `prefers-reduced-motion: reduce` none runs from our CSS. The behavior of Fluent's own Skeleton and Spinner under reduced motion is recorded (Q18). | Mocks (grep for animation/transition in our styles), Visual |

### Strings, icons, manifest, build
| ID | Criterion | Verify by |
|---|---|---|
| F-012/AC-23 | Every user-visible string comes from `strings.ts`; no text literal in the new TSX other than technical values (a test or grep with an explicit allow-list). `html lang` matches the chosen language. With D6 = PT-PT, the table values are used and the user has reviewed the wording for European Portuguese. | Mocks (grep or test), user review |
| F-012/AC-24 | `assets/icon-16.png`, `32`, `64`, `80`, `128` are replaced: file names unchanged, exact pixel sizes (checked by reading the PNG header), 32-bit RGBA, the brand design of this spec (or the user's choice from Q3), an SVG master is in the repo, and a person has looked at 16 and 32 px on a light and a dark background and finds them legible. | Mocks (size test) and Visual |
| F-012/AC-25 | Live: the ribbon button and File > Manage add-ins show the new icons in classic Outlook. Record what was needed to refresh a cached icon (reload, clear cache, restart) and the build. | Live (user) |
| F-012/AC-26 | `manifest.xml` is byte-identical to its state before F-012 (so no admin re-upload is needed), `manifest.test.ts` and `npm run validate` pass, and `dist/manifest.xml` still has `ReadItem` only. If the user decides in Q13 to change manifest text now, this AC becomes: version bumped, the change listed, the admin re-upload planned. | Mocks (diff, command output) |
| F-012/AC-27 | No new network host: no webfont, CDN stylesheet, external image or analytics. The production bundle still passes F-011 checks (hashed names, no source maps, no secrets, no `localhost` in the manifest). Prototype files are not in `dist/`. Bundle size before and after is recorded and the growth stays inside the budget set in plan step 0 (default proposal: at most +50 KiB gzipped over the F-011 baseline). | Mocks (build output, listing, grep) |
| F-012/AC-28 | The only dependency change is `@fluentui/react-icons` as a direct dependency at the exact version already in the lockfile (2.0.344). `package-lock.json` diff shows nothing else new. `npm audit` output is recorded, with no new High or Critical. (Accessibility test tooling, if the user wants it, is Q14 and needs a separate yes.) | Mocks (diff, command output) |
| F-012/AC-29 | Typecheck, lint, tests, `npm run build`, `npm run build:dev` and `npm run validate` pass. After the merge, the F-011 workflow deploys and a code-only update reaches the already deployed add-in without a manifest re-upload (as F-011/AC-17). | Command output, Live (user) |
| F-012/AC-30 | Visual acceptance by the user in classic Outlook on Windows, in light and dark, at the default pane width and one narrower and one wider, for: message loaded, loading, no message, body error, and each of the six backend states that can be provoked (connected, needSignIn, unsupported, notConfigured are reachable; error by blocking the network). Screenshots, if kept, must not contain real email content (CLAUDE.md rule 5). | Live (user), recorded in the plan |
| F-012/AC-31 | Process gate: before implementation the user has chosen a direction and an accent from the throwaway prototype (Q1, Q2), and the prototype code is not merged. | Plan evidence |

## Edge cases
- **Very long unbroken subject, address or body line** (URLs, tokens): wraps; no horizontal scroll (AC-17).
- **Sender without a display name:** the address only. **Both empty:** show the "De" row with an em dash or omit it; never `<>`.
- **Right-to-left or mixed-direction mail:** `dir="auto"` on email-derived text.
- **Emoji, accents, CJK:** displayed literally (existing tests kept).
- **Pane narrower than 320 px or wider than 500 px:** below 320 the layout may scroll horizontally only for the body preview's own scroll area; above 500 the column is capped at 640 px.
- **Windows display scaling and 200% text:** AC-18.
- **Theme changes while the pane is open:** handled if the event fires (AC-7, AC-8); if the client does not fire it, the pane follows the theme on next open (to record).
- **Outlook client without Mailbox 1.14** (for example Outlook 2021 volume-licensed): `matchMedia`, then light. These clients cannot sign in anyway (F-002).
- **Backend slow or down:** the message section works independently of the footer (F-002 rule, kept).
- **Item changes while the pane is open:** F-001 and F-002 do not handle `ItemChanged`. The redesign does not add it, so a pinned pane can show a stale message. This is an existing limitation that becomes more visible with a polished UI; it is a candidate follow-up, not part of F-012.
- **High contrast or forced colors:** the high-contrast theme is used and custom colors are dropped.
- **Webview too old for Fluent v9:** unchanged from today (D8). Not addressed here.
- **Pre-ready flash:** the container is empty until `Office.onReady`; only the page background follows `prefers-color-scheme`.

## Out of scope
- Any AI feature, backend call, prompt, or the data behind the future sections (F-003 onward). The "Detalhes extraídos", "Traduzir", "Verificar tom", "Seguimento", "Competências" tools are not in `requirements.md`; adding them is a product decision.
- "Atividade IA" as a real log (it would store or hold per-action history). Only the static assurance line.
- New permissions, Graph, `Mail.ReadWrite` or drafts (F-009).
- Compose mode and the `Escrever` view (F-010; the manifest has only the read surface).
- Localization machinery (language switching by `displayLanguage`, manifest `Override`s). Strings stay in `strings.ts`.
- A first-run experience or onboarding carousel (there is nothing to configure; the pane is usable on first open).
- A pane settings menu, user theme override, or any persisted preference.
- `ItemChanged` handling or pinning support.
- Changing `manifest.xml` (see Manifest review), the production manifest `<Id>` question from F-011, and a custom domain.
- Webfonts (the Rational font is not licensed for web use here), analytics, or any telemetry.
- Changing `auth.ts`, `backend.ts`, `config.ts` or the Supabase function.

## Risks
- **Brandbook not read.** The PDF could not be rendered. Mark usage, minimum sizes, inverse wordmark and font rules are unconfirmed (Q15). The website's own tokens and favicon are the only evidence.
- **Mockup not seen directly.** The author worked from a description. The prototype is where the layout is compared with the image.
- **Orange is off-brand and low contrast.** If chosen, use the darker orange (white label 5.18:1), not the mockup-like orange (3.04:1).
- **Theme detection is under-documented for Outlook.** Microsoft documents the API and event, not what values Outlook returns per theme, nor whether the reading pane's dark mode follows the Office theme. The design falls back safely, but AC-8 must be run live.
- **Pane title duplication.** Microsoft says not to repeat the pane name in an adjacent top bar, and also says Outlook add-ins skip the footer brand bar. Whether Outlook already shows "DIMSCALE Mail AI" above the pane is not documented here (Q11). If it does, the header drops the text and keeps a small mark.
- **Cached icons.** Replacing PNGs at the same URL may leave old icons in Outlook's cache (F-011 cites the clear-cache page). AC-25 records it. Changing the file names would avoid it but needs a manifest change.
- **Scope creep.** The mockup shows seven tools, a log and an extraction feature that are not in the backlog. The slot table keeps them as hidden, specless placeholders and asks the user to decide (Q7, Q8).
- **Wording that overpromises.** "Criar rascunho" and "Resposta criada em Rascunhos" are only true with F-009, or if a live test proves `displayReplyFormAsync` leaves a draft. The spec uses "Abrir resposta para revisão".
- **Bundle growth.** Fluent components and icons add weight, and the build already warns on size. AC-27 sets a budget and records the delta.
- **Existing tests assert exact strings and DOM shapes** (for example `pre` text length, "From:" text). Migrating them is the largest mechanical risk to AC-1.

## References
- https://learn.microsoft.com/office/dev/add-ins/design/add-in-design
- https://learn.microsoft.com/office/dev/add-ins/design/task-pane-add-ins
- https://learn.microsoft.com/office/dev/add-ins/design/branding-patterns
- https://learn.microsoft.com/office/dev/add-ins/design/first-run-experience-patterns
- https://learn.microsoft.com/office/dev/add-ins/design/accessibility-guidelines
- https://learn.microsoft.com/office/dev/add-ins/design/add-in-icons
- https://learn.microsoft.com/office/dev/add-ins/design/microsoft-365-extension-management-icons (unified-manifest app icons, 192 x 192; not applicable to the XML manifest, read to confirm that)
- https://learn.microsoft.com/javascript/api/manifest/iconurl
- https://learn.microsoft.com/javascript/api/manifest/highresolutioniconurl
- https://learn.microsoft.com/office/dev/add-ins/develop/localization
- https://learn.microsoft.com/office/dev/add-ins/quickstarts/fluent-react-quickstart
- https://learn.microsoft.com/javascript/api/office/office.context (`officeTheme`, `displayLanguage`)
- https://learn.microsoft.com/javascript/api/office/office.officetheme
- https://learn.microsoft.com/javascript/api/outlook/office.mailbox (`addHandlerAsync`, `OfficeThemeChanged`)
- https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-requirement-set-1-14
- https://learn.microsoft.com/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets
- https://learn.microsoft.com/javascript/api/outlook/office.messageread (`displayReplyFormAsync`)
- https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview
- https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache
- https://react.fluentui.dev (Fluent UI React v9; theming and component pages not fetched, API checked in the installed package)
- WCAG 2.1 contrast minimum: https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html, non-text contrast: https://www.w3.org/WAI/WCAG21/Understanding/non-text-contrast.html (formula applied; pages not fetched in this session)
- Brand inputs: `dimscale.newsite/src/app/globals.css`, `src/app/icon.svg`, `public/brand/dimscale-wordmark.svg`.

## Open questions
Each has a recommendation. Questions 1 to 6 block plan step 2 (implementation); the rest can be settled during the prototype.

1. **Direction.** A "Referência" (recommended), B "Navy header" or C "Cards"? The prototype shows all three.
2. **Accent color.** A-navy (brand, recommended; teal on dark) or A-orange (as in the mockup, with the darker orange for any surface with text)? Please send the mockup's orange hex if A-orange.
3. **Add-in icon.** Use the brand navy tile with the white "D" (recommended; same as the website favicon), with or without the teal sparkle? Or the green "D" on black you sent? (Green on black is not a DIMSCALE color.)
4. **UI language (D6).** PT-PT (recommended, matches the mockup) or EN? With PT-PT the literal strings in F-001 and F-002 ACs change (AC-1 amendment) and `html lang` becomes `pt-PT`. Do you approve that amendment?
5. **Production visibility of unbuilt sections.** Hide them and show one "chegam em breve" card (recommended), or show all sections as "Disponível em breve" placeholders?
6. **Tab bar.** Render it only when two or more views exist (recommended; so none in F-012), or show Analisar, Responder and a disabled Escrever now? Also: in the future, which sections belong to which tab? The default in this spec is that Analisar holds all sections as in the mockup, and each feature spec may move its section to Responder.
7. **Quick tools.** Which tools do you want, and what is "Competências" (skills, competencies, something else)? "Seguimento" (follow-up) should remain a suggestion text; reminders or message flags would need Graph write access.
8. **"Detalhes extraídos".** Is data extraction (deadline, project, required action) a feature you want? If yes it needs an ID (proposed F-014), a `requirements.md` entry and its own limit on what text goes to the AI.
9. **Reply button wording (F-005 versus F-009).** Confirm "Abrir resposta para revisão". Whether `displayReplyFormAsync` leaves a draft is not stated in Microsoft's page, so F-005 must test it live before any "rascunho" wording is used. Graph drafts (F-009) are blocked while `Mail.ReadWrite` is revoked.
10. **Font.** Segoe UI for body with the brand stack for headings (recommended), or the brand stack everywhere (renders as Arial on Windows)? Is a licensed Rational web font planned?
11. **Pane title.** Does classic Outlook already show "DIMSCALE Mail AI" above the pane? To be seen live; if yes the header keeps the mark only.
12. **Theme and contrast behavior in Outlook.** What does classic Outlook return in `bodyBackgroundColor` per Office theme, does `OfficeThemeChanged` fire, and does the webview honor `forced-colors`? Answered by AC-8 and AC-9, not by the docs. Should the pane use Outlook's own background color to avoid a seam, instead of the DIMSCALE navy-deep (recommended: DIMSCALE colors, then judge the seam live)?
13. **Manifest text.** Keep it unchanged in F-012 (recommended, no admin re-upload) and bundle the new text with the first AI feature, or change it now?
14. **Accessibility tooling.** Add `axe-core` as a dev-only dependency for automated checks (Microsoft recommends Accessibility Insights in CI)? Not assumed; it needs your yes.
15. **Brandbook.** Please check the PDF for: wordmark minimum size, clear space, inverse (white) version, allowed use of the "D" mark alone, teal usage, and the Rational font license.
16. **Message section default.** Expanded (recommended while there is nothing else) and collapsed by default once AI sections exist?
17. **Pane width.** Microsoft gives 320x570 for Outlook on the web and nothing for Outlook on Windows. Is 320 to 500 px right for DIMSCALE's users? Measure live.
18. **Reduced motion.** Whether Fluent's Skeleton and Spinner honor `prefers-reduced-motion` is not confirmed; record it (AC-22).
19. **Fluent docs.** The Fluent v9 theming docs page returned no text to the fetch tool. The theme API was verified from the installed package's type definitions and token source (9.74.10) instead. A reviewer with a browser should confirm the `BrandVariants` ramp guidance on react.fluentui.dev.

## Brandbook rules (read from the DIMSCALE Brand Book, 2026-10-22)
The source files are in the local, git-ignored `Modelo_DIMSCALE/` folder; the brandbook and the `.ai` file are never committed to this public repo.
- **Logo colour:** the logo is used only in **black or white**, on clean surfaces. It is never coloured, and never placed in colour over a colour (p. 7, "WRONG 3"). On complex backgrounds it sits in a box with safety margins.
  - This means: black wordmark on light theme, white wordmark on dark theme or on a navy band.
  - It also rules out a teal- or orange-tinted logo.
- **Minimum size:** 20 mm for the digital wordmark only (p. 4), about 76 px. The clear space equals the logo height `x` (p. 3).
- **Typography:** the main family is **Rational** (René Bieder, 2016; commercial). For **"e-mail or other online editable platforms"** the brandbook prescribes **Helvetica** (and Courier). Flama is used for "ARCH COST MGMT" and text highlights.
  - The task pane therefore uses the Helvetica stack; no font files are bundled. The PDF embeds only subset or DEMO fonts, which are not reusable.
- **Icons (16–128 px):** the full wordmark is unreadable at these sizes, so the proposal is the "D" glyph from the wordmark, **white on a black square**, which keeps to the black-and-white rule. The user's green-"D" sample breaks the colour rule and is listed as an open question.
- **Tagline** "Architecture Cost Management ®": its height must match the logo height when it sits beside it. It is omitted in the pane header for space.

