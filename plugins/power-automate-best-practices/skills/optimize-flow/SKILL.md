---
name: optimize-flow
description: Review an existing Power Automate cloud flow for performance anti-patterns (unnecessary loops, per-item reads or writes, unfiltered reads, missing pagination, concurrency races, terminate-instead-of-trigger-condition, late responses) and rewrite it with proven patterns. Use when the user asks to optimize, speed up, review, audit or reduce the cost or API calls of a flow. Works on a flow in an environment (via the power-automate plugin's FlowAgent tools) or on a definition JSON file.
user-invocable: true
argument-hint: "<flow name, flow id, or path to a definition JSON>"
---

# Optimize a cloud flow

Target: the user's flow, given as `$ARGUMENTS` (a name, a flow id, or a file path).

This skill uses the guidance, patterns and linter of the sibling skill **flow-best-practices**:

- Linter: `../flow-best-practices/scripts/lint-flow.mjs` (relative to this skill's directory)
- Fix patterns: `../flow-best-practices/references/patterns.md`
- Limits and measurements: `../flow-best-practices/references/limits.md`

Read `../flow-best-practices/SKILL.md` first if it is not already loaded.

## Workflow

1. **Get the definition.**
   - From a file: use it as is.
   - From an environment, with the power-automate plugin's FlowAgent tools (`mcp__flowagent__<tool>` in Claude Code,
     `flowagent-<tool>` in Copilot CLI): `get_current_env` / `list_environments`, `list_flows` (filter by name), then `get_flow`.
     Save the returned definition to a temporary JSON file. If the FlowAgent tools are missing, ask the user to install the
     power-automate plugin, or to export the flow (solution export, or the designer's code view) and give you the file.

2. **Lint it.** Run `node <flow-best-practices dir>/scripts/lint-flow.mjs <file>` (Markdown) and, if you need to process the
   result, `--format json`. The linter is deterministic: report what it found, don't invent findings it didn't make.

3. **Look at real runs when available.** `get_run_history`, then `get_run_actions` on a recent slow run: which action or
   loop owns the run time, how many iterations, any retries or 429s (`get_run_action_repetitions` for loops). Use this to rank
   findings by real impact. If the flow has no runs, say so; don't guess durations.

4. **Rank and explain.** For each finding: what it costs (actions per iteration × items, API calls, payload, correctness risk),
   the fix pattern (section of `patterns.md`) and what changes for the user. Errors first (wrong results: races, truncated reads),
   then the warnings with the biggest volume, then info.

5. **Propose, then ask.** Show the planned changes action by action and ask before changing the flow. Some fixes change behaviour
   (for example: SharePoint `$batch` is not atomic; trigger conditions mean irrelevant events leave no run history). Say so.

6. **Apply.**
   - In an environment: prefer the power-automate plugin's `edit_flow` for surgical changes; use `preview_update` to show the
     diff; `get_backup` / `restore_backup` exist if you need to roll back. Always run `validate_flow` and `preflight_flow`
     before saving, exactly as the power-automate plugin requires. Never bypass validation.
   - On a file: write the optimized definition next to the original (`<name>.optimized.json`).
   - Re-run the linter on the result. It must have no errors, and any warning left must be explained.

7. **Measure.** Ask the user to run the flow with the same input (or use `run_flow` / `smoke_test` if it is safe to run),
   then compare run durations and action counts with the earlier runs. Report measured numbers only; label estimates as estimates.

## Rules

- Never change what the flow does without saying so: same inputs, same outputs, same side effects, unless the user agrees.
- Keep the power-automate plugin's definition rules: `OpenApiConnection`, slash-path parameters, no `authentication` block on
  connector actions, no `Invoker` connection references, no legacy `ApiConnection`.
- Don't raise concurrency on a loop that writes to shared variables or where order matters.
- Respect target limits when you add concurrency or batching: SharePoint 600 calls/min per connection and ≤ 500 items per
  `$batch`; Dataverse 52 concurrent requests and 6,000 requests per 5 minutes per user.
- If the user insists on keeping a pattern (for example one email per recipient), keep it and make the loop thin and concurrent instead.

## Output

End with a short summary table: finding → change made → expected or measured effect, plus anything the user must do
(bind a connection for HTTP with Microsoft Entra ID, index a SharePoint column, confirm a pagination threshold).
