---
name: flow-best-practices
description: Performance and throughput best practices for Power Automate cloud flows. Use whenever you design, generate, edit or review a cloud flow definition (including when the power-automate plugin's build-flow, create-flow or edit_flow is used), so the flow avoids unnecessary loops, per-item API calls, unfiltered reads, silent pagination truncation, concurrency races and runs that should never start. Includes a deterministic linter (scripts/lint-flow.mjs) to run before validate_flow.
user-invocable: true
argument-hint: "[path to a flow definition JSON to check]"
---

# Power Automate flow best practices

Every action run in a cloud flow is billed, throttled and timed the same way, whether it does useful work or not.
A loop over 5,000 rows with 3 actions inside is 15,000 action runs: more than a third of a Premium user's daily
allowance (40,000 Power Platform requests), for one run. Measured in our lab, the patterns below made the same flows
12–177× faster with an identical result (see `references/limits.md`).

## How this fits with the power-automate plugin

The `power-automate` plugin (microsoft/power-platform-skills) makes a flow **valid**: exact connector operations,
`OpenApiConnection` shapes, connection references, `validate_flow` / `preflight_flow`. This skill makes it **efficient**.
Follow both. When they overlap, the power-automate plugin's rules win for definition shape:

- `OpenApiConnection` with `host.apiId` / `connectionName` / `operationId`, slash-path parameter keys (`body/…`, `parameters/…`).
- **No `authentication` block** on connector actions (the snippets in `references/patterns.md` follow this).
- Always call `get_operation_details` for exact parameter names; never guess.

Order of work when building or changing a flow:

1. **Design** with the decision framework below, before writing any JSON.
2. **Generate** the definition using `references/patterns.md`.
3. **Lint**: save the definition to a file and run `node <this skill's directory>/scripts/lint-flow.mjs <file> --fail-on warning`.
   Fix every error. Fix every warning, or tell the user why it stays (for example, "one email per recipient is a real side effect").
4. **Validate** with the power-automate plugin: `validate_flow`, then `preflight_flow`.
5. **Create** the flow stopped (`create_flow`), as the power-automate plugin says, and tell the user what you optimized.

## Decision framework: do you need a loop at all?

Ask in this order. Most loops disappear at question 1 or 2.

| # | Question | If yes |
|---|---|---|
| 1 | Can the **source** do the work? | Push it down: `$filter`, `$select`, `$expand`, `$top`, FetchXML aggregates, trigger conditions |
| 2 | Is it only **reshaping, filtering, joining or aggregating** data? | No loop: Filter array, Select, a dictionary object, `xml()` + `xpath()`, `union()`, `sort()`, `length()`, `max()` |
| 3 | Does the target accept **many records per call**? | Batch: loop over `chunk(items, N)`, not items. SharePoint `$batch` (N ≤ 500), Dataverse `CreateMultiple` / `UpdateMultiple` / `UpsertMultiple` (N = 100–1,000), Excel Office Script `setValues()` |
| 4 | Still need **one side effect per item** (an email per person, an API with no batch)? | Loop, but thin (few actions inside) and concurrent (20–50) unless order matters |

## Non-negotiable rules

1. **No shared variables inside a concurrent loop.** `Set variable` with concurrency > 1 is a race and gives wrong results.
   Use `item()` or a Compose inside the loop, or a Select after it. Append/Increment are locked, so concurrency buys nothing.
2. **Every list read declares how much it reads.** SharePoint *Get items* returns only **100 items** unless pagination is on
   (`runtimeConfiguration.paginationPolicy.minimumItemCount`) or `$top` is set. Dataverse *List rows* stops at 5,000.
   Size the threshold to the data. Don't use 100,000 "just in case".
3. **Filter and trim at the source.** Every read has a `$filter` (or FetchXML) and a column list (`$select`, a view). Index the
   SharePoint columns you filter on for lists over 5,000 items. No Filter array straight after a read when OData can do it.
4. **No read or write per item when a bulk path exists.** No *Get item* per row: read once and join with a dictionary, or use
   `$expand`. No *Create item* / *Add a new row* / *Add a row into a table* per row: batch it.
5. **SharePoint `$batch`: ≤ 500 items per request, and check every part.** SharePoint stops a batch after about 90 s of
   server time and returns HTTP 200 anyway. Changesets are **not** atomic: count `HTTP/1.1 2xx`, `4xx` and `5xx` parts.
   Dataverse changesets and `CreateMultiple` are all-or-nothing.
6. **Automated flows filter with trigger conditions**, not with a Condition + Terminate inside the flow: those runs are still billed.
   Don't re-read the item that triggered the flow; the trigger output already has its fields.
7. **Callers that wait get an answer within 120 s.** Power Apps and HTTP callers fail at 120 s and the run is marked Failed.
   For anything that can run longer, respond first (status + run id), then work, and report progress through Dataverse or a list.
8. **Compose for values set once; variables only for real accumulators** in sequential logic.
9. **Do until has explicit limits** (`limit.count`, `limit.timeout`). The defaults (60 iterations, 1 hour) stop silently.

## Platform tuning

| Setting | Guidance |
|---|---|
| Apply to each concurrency | Default 1. 20–50 for independent per-item work. Respect target limits: SharePoint 600 calls/min per connection, Dataverse 52 concurrent requests and 6,000 requests / 5 min per user |
| Trigger concurrency | Limit parallel runs (1–100) to protect a downstream system or keep order. It can't be turned off again on that flow |
| Parallel branches | Independent reads side by side: total time = the slowest branch |
| Fan-out | For very large jobs, `chunk()` + Apply to each (concurrency 10) + *Run a Child Flow*. Saves wall-clock time, not actions |
| Retry policy | Honour 429 + `Retry-After`; keep batch calls idempotent |
| Scope + `result()` | Wrap the work in a Scope and read errors with `result('Scope')`, so failures inside loops are reported (a Scope whose last action succeeds is *Succeeded* even if a loop inside failed) |

## Expression gotchas

- `xml()` needs one root and valid element names (no spaces, no leading digit): rename with Select first.
- `xpath(... 'sum(...)')` returns a float. Format it with `formatNumber()`.
- `add()` takes exactly two arguments. Nest it.
- `result('<loop>')` returns one entry per action, whose `outputs` is an array of every iteration.
- Inside a SharePoint `$batch`, use `PATCH` in the request line (`X-HTTP-Method: MERGE` is ignored there).
- A SharePoint lookup column `Customer` is `CustomerId` in REST.

## Report to the user

After building or changing a flow, include a short best-practice summary:

- Loops: how many, actions per iteration, concurrency.
- Linter result: errors, warnings, info, and what you fixed or deliberately kept (with the reason).
- Assumptions about data volume (pagination threshold, batch size) that the user should confirm.

Never claim a speed-up you didn't measure. Suggest measuring it: run history action durations before and after, or
the Dataverse `flowrun` table (duration in milliseconds).

## References

- `references/patterns.md`: copy-ready JSON for every fix (Filter array + Select, dictionary join, xpath aggregation,
  source filtering + pagination, SharePoint `$batch`, Dataverse `CreateMultiple`, Office Script, trigger conditions,
  respond early, fan-out).
- `references/limits.md`: platform limits and the lab measurements behind these rules.
- `scripts/lint-flow.mjs`: the linter (rules BP001–BP016). `--format json` for machine-readable output.
