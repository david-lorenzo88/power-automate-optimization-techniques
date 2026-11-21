# Power Automate best practices plugin

Skills for **GitHub Copilot** and **Claude Code** that make AI-built Power Automate cloud flows fast and cheap to run,
not just valid. It complements the [`power-automate` plugin](https://github.com/microsoft/power-platform-skills/tree/main/plugins/power-automate)
from microsoft/power-platform-skills:

| | `power-automate` (Microsoft) | `power-automate-best-practices` (this plugin) |
|---|---|---|
| Focus | Correctness: connectors, connection references, validation, deploy, runs | Performance: loops, bulk APIs, source filtering, pagination, concurrency, trigger conditions |
| Tools | FlowAgent MCP (`build-flow`, `edit_flow`, `validate_flow`, …) | Guidance + patterns + a deterministic linter (`lint-flow.mjs`) |
| Together | The agent builds with FlowAgent, applies these rules, lints, then runs `validate_flow` / `preflight_flow` | |

Every rule and pattern comes from flows that were built and measured: the same input and the same result, 12–177× faster
([lab results](../../lab/RESULTS.md)).

## Skills

| Skill | When it's used |
|---|---|
| `flow-best-practices` | Automatically when the assistant designs, generates, edits or reviews a cloud flow. Also `/flow-best-practices <file>` to check a definition |
| `optimize-flow` | `/optimize-flow <flow name, id or file>`: audits an existing flow, ranks findings by impact, proposes fixes and applies them with FlowAgent after you confirm |

## Install

Install both plugins: Microsoft's `power-automate` plugin for building and deploying, and this one for the best practices.

**Claude Code or GitHub Copilot CLI**, inside a session:

```bash
/plugin marketplace add microsoft/power-platform-skills
```

```bash
/plugin install power-automate@power-platform-skills
```

```bash
/plugin marketplace add david-lorenzo88/power-automate-optimization-techniques
```

```bash
/plugin install power-automate-best-practices@power-automate-optimization
```

**Without the plugin system** (VS Code with GitHub Copilot, the Copilot coding agent, or a single project): copy the two
skill folders into the repository where you build flows. Keep them side by side, because `optimize-flow` uses the linter
and references in `flow-best-practices`.

| Host | Copy `skills/flow-best-practices` and `skills/optimize-flow` to |
|---|---|
| GitHub Copilot (VS Code, coding agent, CLI) | `.github/skills/` |
| Claude Code | `.claude/skills/` |

Requires Node.js 18+ for the linter. No npm install.

## Try it

- "Build a flow that copies new SharePoint orders over 1,000 EUR into Dataverse every night"
- "Optimize my flow *Nightly order sync*, it takes two hours"
- `/flow-best-practices ./exports/MyFlow.json`

## The linter

```bash
node skills/flow-best-practices/scripts/lint-flow.mjs <flow.json> [--format md|json] [--fail-on error|warning|none]
```

Accepts a raw workflow definition, Dataverse `clientdata` (object or string), or a FlowAgent `get_flow` result.
Exit code 1 when a finding is at or above `--fail-on` (default `error`), so it also works as a CI gate on exported solutions.

| Rule | Severity | Finds |
|---|---|---|
| BP001 | warning | Loop that only reshapes, filters or aggregates data |
| BP002 | warning | Nested loop |
| BP003 | warning | Read call inside a loop (Get item / query per row) |
| BP004 | warning | One write per item (Create item, Add a new row, Add a row into a table, …) when a bulk path exists |
| BP005 | error | Shared variable written inside a concurrent loop (race) |
| BP006 | info | Sequential loop with independent side effects |
| BP007 | warning | SharePoint / Dataverse read without a filter |
| BP008 | info | Read that returns every column |
| BP009 | error | SharePoint Get items / Dataverse List rows without pagination or `$top` (silently truncated) |
| BP010 | warning | Filter array right after a read |
| BP011 | warning | Automated trigger without trigger conditions, followed by Condition + Terminate |
| BP012 | info | Get item right after the trigger that already returned it |
| BP013 | warning | Response after loops / delays / child flows (120 s limit) |
| BP014 | info | Variable that is never changed (use Compose) |
| BP015 | warning | SharePoint `$batch` chunks larger than 500 |
| BP016 | info | Do until without explicit limits |

Tests: `npm test` in this folder. They lint the lab's SLOW flows (each must be flagged with its anti-pattern) and FAST flows
(no errors or warnings), plus synthetic definitions for every rule.
