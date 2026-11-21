# AGENTS.md — power-automate-best-practices plugin

Guidance for AI agents (GitHub Copilot, Claude Code) using this plugin.

## Routing

- Building, generating, editing or reviewing a Power Automate **cloud flow** → load `skills/flow-best-practices/SKILL.md`
  and apply it on top of the power-automate plugin's workflow.
- The user asks to optimize, speed up, audit or cut the cost of an existing flow → `skills/optimize-flow/SKILL.md`.
- Desktop flows (RPA), canvas apps, Power Pages: out of scope.

## Working with the power-automate plugin (microsoft/power-platform-skills)

- Use its FlowAgent MCP tools (`mcp__flowagent__*` / `flowagent-*`) for every environment, flow, connection and run operation.
- Its definition rules win on shape: `OpenApiConnection`, slash-path parameter keys, no `authentication` on connector
  actions, `Embedded` connection references, `get_operation_details` before every connector action.
- Run `scripts/lint-flow.mjs` **before** its `validate_flow` / `preflight_flow`. Lint findings never replace validation.

## Non-negotiable

- Don't report a speed-up you didn't measure. Label estimates as estimates.
- Don't change a flow's behaviour (inputs, outputs, side effects, atomicity) without telling the user first.
- Don't add concurrency to loops that write shared variables or depend on order.
- The linter's findings are the source of truth for "what is wrong". Don't invent findings it didn't report; you may add
  observations from run history, labelled as such.
