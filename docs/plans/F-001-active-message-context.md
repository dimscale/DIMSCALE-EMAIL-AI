# Plan F-001: Task pane shell and active message context

- **Spec:** [F-001](../specs/F-001-active-message-context.md). **Status:** In progress.

| # | Step | ACs | Agent |
|---|---|---|---|
| 1 | Run `yo office` (Outlook, task pane, TypeScript, React if offered, XML manifest) in a temp folder, then move the output into the repo root without overwriting `CLAUDE.md`, `docs/`, `.claude/` or `.gitignore`. Run `npm run build` and `npm run validate`. | AC-1 | coder |
| 2 | Restrict the manifest to `ReadItem` and the message read surface. Remove sample code the spec doesn't need. | AC-2 | coder |
| 3 | Add a pure function `toMessageView(item, bodyText)` that applies the subject, sender and truncation rules, with unit tests. | AC-3, AC-4, AC-5 | coder |
| 4 | Panel component: loading, empty item, error with Retry, and success states. Text-only rendering. Tests with a mocked `Office.context.mailbox.item`. | AC-3, AC-6, AC-7, AC-8 | coder |
| 5 | Run typecheck, lint, tests and build, and record the output below. | AC-10 | qa-tester |
| 6 | Independent review. | AC-8, AC-9 | code-reviewer |
| 7 | Manual sideload in Outlook on the web or new Outlook (user-assisted), with the result recorded below. | AC-3 (live) | user + orchestrator |

## Evidence

### 2026-10-22: steps 1 to 6 (re-run by the orchestrator after the fixes)
- **Scaffold:** `yo office` React + TypeScript task pane, Outlook, XML manifest. Generator samples were removed, along with `manifest.json` and `@fluentui/react-icons`.
- **Checks:**

| Check | Result |
|---|---|
| `npm run typecheck` | exit 0 (`strict: true`) |
| `npm run lint` | exit 0 |
| `npm run test` | 4 files, 31 tests passed, 0 `it.fails` (mocked Office.js) |
| `npm run build:dev` | "compiled successfully" |
| `npm run build` | exit 2, fails on purpose: "Production build blocked: set urlProd…" (D7) |
| `npm run validate` | "The manifest is valid." (Outlook web, Windows, Mac) |
| grep `src/` (non-test) for `console.`, `innerHTML`, `fetch(`, storage | no matches |

- **QA (qa-tester, haiku):** AC-1, AC-2, AC-4, AC-5, AC-7, AC-8, AC-9 and AC-10 pass (mocked or static). AC-6 failed on a synchronous `getAsync` throw (D1); fixed and now tested. AC-3 is mocked only.
- **Review (code-reviewer, sonnet): APPROVE WITH COMMENTS**, with no Critical or High findings. Medium items fixed:
  - the production-URL placeholder now blocks the production build;
  - the date field is relabelled "Date" (Office.js `MessageRead` has only `dateTimeCreated`, verified on Microsoft Learn);
  - `strict` TypeScript is on.

  Low items fixed: surrogate-pair truncation (D2), whitespace-only body (D5), and the unused icons dependency. Low items accepted:
  - IE11 polyfills (status D8);
  - `bt:Sets DefaultMinVersion="1.3"` (needed for add-in commands);
  - unused devDependencies left by the generator;
  - an empty "From:" when the sender has neither name nor address.
- **Step 7 (live sideload):** NOT DONE. It needs the user: run `npm start`, accept the dev certificate, sign in to M365, then open a message → Apps → DIMSCALE Mail AI. Run the QA manual tests 1 to 9.
