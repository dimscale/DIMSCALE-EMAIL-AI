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
_(fill in with real command output and dates)_
