# Test Strategy

## Levels
| Level | What | How |
|---|---|---|
| Static | Types, lint, manifest | `npm run typecheck` (or `tsc --noEmit`), `npm run lint`, `npm run validate` |
| Unit | Pure logic (mapping, truncation, prompt building, output-schema validation) | The generator's test runner, or one added per ADR-0001 |
| Component | Panel states (loading, empty, error, success) with mocked `Office.context.mailbox` | Same runner, with a DOM test environment |
| Backend | Token validation, input limits, prompt delimiting, schema validation, with a mocked OpenAI | Function unit tests |
| Manual / live | Sideload in Outlook (web or new Outlook first) with a test mailbox | Recorded in the plan's Evidence section, with the date, client and result |

Exact commands get filled in here once F-001 has scaffolded the project.

## Mandatory checks per feature
1. Typecheck, lint, tests and build all pass, and the output is pasted in the plan's Evidence section.
2. Every AC maps to a test name or an explicit manual step.
3. Safety tests where relevant:
   - HTML or script in email text renders literally
   - nothing sends, deletes or moves mail
   - nothing logs email content or tokens
   - prompt-injection fixtures (an email saying "ignore previous instructions…") don't change the output schema or trigger actions

## Honesty rules
- Mock-based results are labeled "mocked".
- "Live Outlook verified" requires a recorded manual run.
- Recommended but unexecuted tests are listed separately.
