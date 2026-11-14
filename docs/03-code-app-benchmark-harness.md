# Code App Benchmark Harness: Feasibility & Design

**Goal.** From a Power Apps **Code App** (React + TypeScript), launch every SLOW / MID / FAST flow variant, then show the results on screen: duration, transactions/actions, API calls, payload size, correctness and speed-up. Each run also gets a link to **open the flow run** and to **edit the flow**.

## Verdict: ✅ Feasible, with three design constraints

| # | Constraint (Microsoft Learn, Nov 2026) | Impact | How we handle it |
|---|---|---|---|
| C1 | Code apps can call **only solution-aware instant flows that use the Power Apps trigger** ("When Power Apps calls a flow (V2)"). Scheduled, automated and other instant triggers are not supported. | Flows triggered by SharePoint or Dataverse (UC7, trigger conditions) can't be launched directly. | All benchmark flows use the Power Apps V2 trigger. For UC7 the app launches a **driver** flow that modifies 100 items, and the two automated flows (with and without a trigger condition) log their own runs. |
| C2 | Synchronous request limit is **120 s**. | SLOW variants run for minutes, so the app can't sit waiting on the call. | **Async pattern**: the app passes a `correlationId`, the flow writes a "Running" telemetry row immediately, does the work, then updates the row. The app polls Dataverse. We don't depend on the flow's response at all. Optionally respond early with the run ID. |
| C3 | Flows are added one by one: `pa app add flow --flow-id <id>` generates a typed `<Flow>Service.ts`. You must re-run it when the flow definition changes. | ~20–25 generated services | Every flow has the **same trigger contract** (inputs/outputs), so the generated code is uniform and a small TS registry maps `scenario + variant → Service.Run`. |

Other requirements:
- Environment with Dataverse, and **"Enable code apps"** turned on (PPAC → Environments → Settings → Product → Features).
- End users (us) need **Power Apps Premium** or pay-as-you-go. They also need Dataverse permissions to invoke flows (**App Opener** role or equivalent).
- Tooling: Node LTS, `npm i -g @microsoft/power-apps-cli` (the `pa` command), client library `@microsoft/power-apps` **≥ 1.1.1**. The flow commands exist only in the npm-based `pa` CLI, not in `pac`. The npm CLI was reported as preview in spring 2026, so check its status when we build.
- Code apps have **no Power Platform Git integration**. This repo is the source of truth for the app code; the flows go in a solution (export unpacked into the repo as well).

---

## Architecture

```mermaid
flowchart LR
  subgraph CodeApp["Code App (React + TS)"]
    UI[Scenario cards<br/>SLOW / MID / FAST · volume] --> REG[Flow registry]
    REG -->|Service.Run(correlationId, volume)| FL
    POLL[Poll every 2 s] -->|BenchmarkRunsService.getAll filter correlationId| DV
    DASH[Results + charts + links]
  end

  subgraph PA["Power Automate (one solution)"]
    FL[UCx-SLOW / UCx-FAST flows<br/>Power Apps V2 trigger]
    DRV[UC7 driver flow] --> AUTO[UC7 automated flows<br/>w/o and with trigger condition]
  end

  subgraph Data
    SP[(SharePoint lists<br/>Orders / Customers / Import)]
    DVD[(Dataverse demo tables)]
    DV[(pa_benchmarkrun<br/>telemetry)]
    FR[(flowrun<br/>platform run history)]
  end

  FL --> SP & DVD
  FL -->|create Running → update Completed| DV
  AUTO --> DV
  POLL -.optional cross-check.-> FR
  DASH -->|open run / edit flow| MAKER[make.powerautomate.com]
```

### Run lifecycle

1. The user clicks **Run FAST** on UC4 at volume 1,000. The app generates `correlationId = crypto.randomUUID()`.
2. The app calls `UC4_FASTService.Run({ correlationId, volume: 1000, label })` and doesn't block the UI on it.
3. The flow's first actions are `Compose Start = utcNow()`, then **Add a new row** to `pa_benchmarkrun` (status Running, run links), then optionally *Respond to a Power App or flow* (returns runId + links).
4. The flow does the work inside a `Scope "Work"`.
5. In a `Scope "Finally"` (run after: succeeded / failed / timed out), it computes the KPIs and updates the row (status Succeeded/Failed, duration, counters, checksum).
6. The app polls `pa_benchmarkrun` by `correlationId` and renders live status → final KPIs, comparing against the latest SLOW run of the same scenario/volume.

---

## Uniform flow contract

**Trigger inputs (Power Apps V2)**

| Name | Type | Purpose |
|---|---|---|
| `correlationId` | Text | Key the app polls on |
| `volume` | Number | Records to process (lets us run SLOW small live, big beforehand) |
| `label` | Text | Free text, e.g. "Dubai live demo" |

**Response (optional, if respond-early proves reliable)**: `runId`, `runUrl`, `editUrl`.

**Expressions used in every flow**

```text
Start          : utcNow()
DurationMs     : div(sub(ticks(utcNow()), ticks(outputs('Start'))), 10000)
EnvironmentId  : workflow()?['tags']?['environmentName']
FlowId         : workflow()?['name']
RunId          : workflow()?['run']?['name']
RunUrl         : concat('https://make.powerautomate.com/environments/', <EnvironmentId>, '/flows/', <FlowId>, '/runs/', <RunId>)
EditUrl        : concat('https://make.powerautomate.com/environments/', <EnvironmentId>, '/flows/', <FlowId>, '/edit')   // verify in PoC
PayloadBytes   : length(string(body('Get_items')))
```

> ⚠️ Community sources report that **solution-aware flows sometimes need `/solutions/~preferred/` in the path** (`.../environments/{env}/solutions/~preferred/flows/{id}/runs/{run}`). We validate both forms in the PoC. The app keeps a fallback: it builds the links itself from the environment ID and a flow-ID registry.

---

## Telemetry table `pa_benchmarkrun` (Dataverse)

| Column | Type | Source |
|---|---|---|
| `pa_correlationid` (primary name / alt key) | Text | trigger input |
| `pa_scenario` | Choice/Text (UC1…UC10) | hard-coded in the flow |
| `pa_variant` | Choice (SLOW / MID / FAST) | hard-coded |
| `pa_technique` | Text, e.g. "Select + Filter array" | hard-coded |
| `pa_volume` | Whole number | input |
| `pa_status` | Choice (Running / Succeeded / Failed) | |
| `pa_starttime`, `pa_endtime` | Date/time | `utcNow()` |
| `pa_durationms` | Whole number | flow-measured |
| `pa_recordsprocessed` | Whole number | `length(...)` |
| `pa_recordsretrieved` | Whole number | items returned by the source (UC6: retrieved vs used) |
| `pa_loopiterations` | Whole number | `length()` of the array looped over (0 for FAST) |
| `pa_connectorcalls` | Whole number | calls to SP/DV/HTTP, computed per design (e.g. N vs `ceil(N/1000)`) |
| `pa_actionsestimated` | Whole number | per-flow formula, e.g. `5 + 3*N`. Doubles as the **Power Platform requests** estimate |
| `pa_payloadbytes` | Whole number | `length(string(body(...)))` |
| `pa_resultchecksum` | Text | e.g. `concat(count,'|',sum)`, so slow and fast can be shown to match |
| `pa_runid`, `pa_flowid`, `pa_environmentid` | Text | `workflow()` |
| `pa_runurl`, `pa_editurl` | URL | computed |
| `pa_error` | Multiline | `result('Work')` filtered on Failed |

The **app computes**: speed-up factor (SLOW ÷ FAST duration), actions saved, % of the 40,000 PPR/day user allowance, a checksum match ✅/❌, and an extrapolation when SLOW ran at a smaller volume (labelled as an extrapolation).

**Platform cross-check (optional):** read the Dataverse **`flowrun`** elastic table (run duration, status, start/end) filtered by `workflowid` / run name. This shows that the self-reported duration matches what the platform recorded. The table must be enabled (28-day default retention), and the flow owner needs read access to FlowRun.

**Stretch (not officially supported):** actual action and repetition counts from the Power Automate management REST API (`.../flows/{id}/runs/{runId}/actions`) via a collector flow. Use it only if time allows; the formula-based estimate is deterministic and easy to explain on stage.

---

## App UI (proposal)

1. **Scenario grid**: one card per UC with the technique name, a volume selector (100 / 500 / 1,000 / 5,000) and buttons **Run SLOW · Run MID · Run FAST · Race (both)**.
2. **Live run panel**: status pill (Running → Succeeded), elapsed timer, then KPI tiles: duration, actions, connector calls, payload, records, checksum match.
3. **Comparison chart** per scenario: SLOW vs FAST duration and actions (log scale option, because "hours vs seconds" doesn't fit a linear axis).
4. **Run history / scoreboard**: every run with links **🔗 Open run** · **✏️ Edit flow**, filters by scenario and label. It's the closing slide of the session, live.
5. **Admin**: Seed / Reset buttons (they call `SEED-*` / `RESET-*` flows).

Generated code usage (from Microsoft Learn):

```ts
import { UC4_FASTService } from './generated/services/UC4_FASTService';
import { Pa_benchmarkrunsService } from './generated/services/Pa_benchmarkrunsService';

const correlationId = crypto.randomUUID();
void UC4_FASTService.Run({ correlationId, volume: 1000, label: 'live' }); // don't await: async pattern

const poll = setInterval(async () => {
  const r = await Pa_benchmarkrunsService.getAll({
    filter: `pa_correlationid eq '${correlationId}'`,
    select: ['pa_status', 'pa_durationms', 'pa_actionsestimated', 'pa_connectorcalls', 'pa_runurl', 'pa_editurl'],
    top: 1,
  });
  // update state; clearInterval when status !== Running
}, 2000);
```

Setup commands:

```bash
pa app init --display-name "Flow Optimization Lab" --environment-id <env-id>
pa app add data-source --connector dataverse --table pa_benchmarkrun
pa app list-flows --search UC
pa app add flow --flow-id <flow-id>     # repeat per flow
npm run build && pa app push
```

---

## PoC checklist (do this before building all 20+ flows)

1. [ ] Turn on code apps in the demo environment and confirm the presenter account has Premium.
2. [ ] Create a solution, the `pa_benchmarkrun` table, and one trivial flow (Power Apps V2 trigger + add row + update row).
3. [ ] `pa app add flow`: confirm the generated service and models, and that `Run()` works.
4. [ ] Test **respond early**: does the flow keep running after *Respond to a Power App or flow*, and does `Run()` resolve at that moment? If not, fire-and-forget + polling is enough.
5. [ ] Test a flow running longer than 120 s: the app must not error out, or the error must be harmless.
6. [ ] Validate both run-URL forms (with and without `/solutions/~preferred/`) and the edit URL. Also check that links open in a new tab from the hosted app (CSP is strict by default in code apps).
7. [ ] Confirm the Dataverse connector's "Perform a bound action" exposes `CreateMultiple` on our table, or use HTTP with Microsoft Entra ID.
8. [ ] SharePoint `$batch`: find the real per-batch limit on our tenant (100 / 500 / 1,000) and whether changeset failures roll back.
9. [ ] `flowrun` table populated for our solution flows → cross-check works.

## Risks & mitigations

| Risk | Mitigation |
|---|---|
| Conference Wi-Fi or tenant throttling during the live demo | Pre-run all SLOW variants at full volume (they stay in the history); record a backup video of each demo |
| The npm `pa` CLI is in preview and syntax changes | Pin the CLI version in `package.json`; keep a README with the exact commands |
| A flow definition changes and the generated service goes stale | The uniform trigger contract means internal changes don't alter the schema; re-run `pa app add flow` anyway before the event |
| SLOW variants take too long live | Use the `volume` parameter: show SLOW@500 live, extrapolate and show the pre-recorded SLOW@5,000 |
| Licensing questions from the audience | One slide: standard vs premium per technique (see catalog) |
