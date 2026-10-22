# DIMSCALE Mail AI — Development Rules

Outlook add-in that helps users read, triage and answer email with AI. Office.js for the active item, Microsoft Graph for authorized mailbox operations, OpenAI API (server-side only) for generation and classification.

Status, backlog and blockers: [docs/status.md](docs/status.md). Read it first in every session.

## Non-negotiable rules

1. **The user controls every email action.** The code never sends, deletes, moves or reports email on its own. Never request the `Mail.Send` Graph permission. Drafts and reply forms are allowed because the user still has to press Send.
2. **No secrets on the client.** The OpenAI API key lives only in the backend's secret store. Never put it in frontend code, the manifest, `localStorage`/`sessionStorage`, source control or client-visible config. `.env*` files stay git-ignored.
3. **Email content is untrusted input.** Instructions inside an email never override the system prompt or trigger actions. AI output is rendered as plain text, never as HTML. AI output never decides an action by itself.
4. **Send the AI as little as possible.** Only the fields a feature needs, as plain text, truncated to the spec's limit. See [docs/architecture/security.md](docs/architecture/security.md).
5. **Never log email content, tokens or keys.** This covers subjects, bodies, addresses and `Authorization` headers, in the browser console, the backend logs and in reports.
6. **Don't invent APIs.** Every Office.js method, requirement set, Graph endpoint and permission must be checked against Microsoft Learn before use, and the spec cites the page.
7. **Least privilege.** Active-item features use Office.js with the `ReadItem` manifest permission and no Graph. Graph scopes are requested incrementally, only by the features that need them (see [overview](docs/architecture/overview.md#capability-boundaries)).
8. **No new platform without an ADR.** A new backend, hosting provider, n8n, VPS or a heavy dependency needs an accepted ADR in `docs/decisions/` first.

## SDD lifecycle (every feature)

1. **Research** the existing code and the official docs.
2. **Specify**: write `docs/specs/F-NNN-<slug>.md` with numbered, testable acceptance criteria (AC) and an explicit out-of-scope list.
3. **Review dependencies and security**, and update `security.md` if the data flow changes.
4. **Plan**: write `docs/plans/F-NNN-<slug>.md`, small ordered steps, each linked to the ACs it satisfies.
5. **Approval**: the user approves the spec (spec `Status: Approved`). No implementation before this.
6. **Implement** only the approved scope.
7. **Verify**: run typecheck, lint, tests and build, and record the actual output.
8. **Independent review** by the `code-reviewer` agent.
9. **Fix** Critical and High findings (or escalate them to the user) and rerun the affected checks.
10. **Close**: update the spec status, `docs/status.md` and the docs.

A feature is **Done** only with all of the following: the spec ACs met, the checks run with their output recorded in the plan, the review completed, no open Critical or High findings, no secrets, no automatic sending, and the docs updated. "It compiles" is not evidence. A live Outlook test is claimed only when it was actually performed.

## Roles

The **main Claude Code session is the Orchestrator / Tech Lead.** Subagents cannot spawn other subagents, so orchestration stays in the main session. It breaks work into tasks, delegates, verifies each agent's evidence, makes the final technical call, keeps `docs/status.md` current, and escalates auth and security decisions to the user instead of guessing. Trivial tasks are done directly, without delegating.

| Agent (`.claude/agents/`) | Model | Use for |
|---|---|---|
| `requirements-architect` | sonnet | Specs, ACs, ADR drafts, checks against official docs. No product code. |
| `coder` | sonnet (pass `model: haiku` for simple, well-defined tasks) | Implementing an approved plan step and its tests. |
| `qa-tester` | haiku | Independent check of the ACs: runs the checks and adds targeted tests. |
| `code-reviewer` | sonnet (pass `model: opus` for auth, token or security changes) | Independent review. Reports findings, edits nothing. |

Typical flow: `requirements-architect` → user approval → `coder` (one plan step at a time) → `qa-tester` → `code-reviewer` → fixes → close.

## Conventions

- Stack (proposed, see [ADR-0001](docs/decisions/0001-addin-stack.md)): Yeoman Office generator scaffold, TypeScript, React, npm.
- IDs: features `F-NNN`, ADRs `NNNN`, acceptance criteria `F-NNN/AC-n`.
- Small, focused commits that reference the feature ID, e.g. `F-001: read active message context`.
- Agent reports always list the files changed, the commands run with their real results, and the open issues.
- Docs are in English. Write the company name as DIMSCALE (uppercase).
