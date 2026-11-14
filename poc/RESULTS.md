# PoC results

Fill in after running the kit (see poc/README.md). ✅ / ❌ / ⚠️ + notes.

**Summary:** #1 ✅ · #2 ✅ · #3 ✅ · #4 ✅ · #5 ⚠️ (respond early or not at all past 120 s) · #6 ⚠️ (use `/flows/…/runs` + `/details`; solution path and `/edit` broken) · #7 ⚠️ (API ✅ 18× faster; not available in *Perform a bound action*) · #8 ⚠️ (≤ 500 items per batch; ~90 s server limit; changesets **not** atomic) · #9 ✅ (`flowrun.duration` in ms)

Run on 2026-11-14 against **PREMIUM-US** (`df93e55c-cf3b-e261-87e1-ddb2b052de1d`, https://premium-us.crm.dynamics.com, Sandbox), as david@powerplatformninja.com.
Scripts and `pa` ran from a Linux cloud container (Node 22.23, Azure CLI 2.92.0, `pa` 1.4.1), signed in with device code.

| # | Check | Result | Notes |
|---|---|---|---|
| 1 | Code apps enabled + Premium | ✅ | Confirmed by the user in PPAC. Preflight: WhoAmI OK (699 ms), `flowrun` readable, prefix `pa` free. |
| 2 | Solution, tables, flows created via script | ✅ | Publisher, solution, `pa_benchmarkrun`, `pa_orderimport` (+ alt key), 2 connection refs, and all 4 flows created **in the solution** via the Web API and turned on. Two setup errors, see *Issues*. |
| 3 | `pa app add flow` + `Run()` works (T3) | ✅ | `pa app init`, `add data-source` and `add flow` ×4 all worked; the app found 1 table + 4 flow services by name. T3: `Run()` returned in 6.9 s (first call, includes warm-up) with the Respond body (`runid`, `runurl`, `respondedat`). Flow itself: 1.9 s. Second session: `Run()` 2.1 s. |
| 4 | Respond early keeps the flow running (T4) | ✅ | Run() resolve time: **1.2 s / 0.8 s** (two sessions). The flow kept running and finished in 60.7 / 60.8 s (flowrun: 61.2 / 61.1 s). |
| 5 | > 120 s: no-respond (T5a) / respond-late (T5b) | ⚠️ | Run() behaviour: **T5a (no Respond)**: `Run()` resolves in **0.5–0.7 s** (async/202) and the flow succeeds in ~150 s, so it's fire-and-forget and the result has to come from `pa_benchmarkrun` polling. **T5b (Respond at end, 150 s)**: `Run()` **rejects at 120 s**. The first session got a generic `{"code":"UnexpectedError"}`; the second got `{"error":{"code":"TimedOut","message":"Invocation of API timed out."}}`. The error shape isn't stable, so the app shouldn't rely on it. The flow's own work completes (`pa_benchmarkrun` = Succeeded, 150.5 s), but **`flowrun.status` = Failed**, because the Respond action fails once nobody is waiting. **Design rule:** any flow that can exceed ~120 s must respond early (or not at all) and report through Dataverse. A late Respond also turns the run red in run history. |
| 6 | Links: flows path / solutions path / edit / details, anchor vs window.open, local vs published | ⚠️ | Published app: **works**: run via `/flows/{id}/runs/{runId}`, `/flows/{id}/details`, run URL built in-app from `getContext()`. **Broken**: run via `/solutions/~preferred/flows/{id}/runs/{runId}`, and `/flows/{id}/edit`. → Use the plain `/flows/` path for run and details links, and drop the solution-path and `/edit` links. The export has one verdict per link, not one per method (anchor vs window.open). Local Play not tested: it serves from `localhost:3000`, which a cloud container can't expose to the user's browser. |
| 7 | CreateMultiple in "Perform a bound action" picker; script 03 timings | ⚠️ API ✅ / picker ❌ | **API**: CreateMultiple 1000 rows in **10.1 s** in 1 request, vs ~182 s extrapolated for sequential single POSTs (~18×). UpdateMultiple 10.7 ms/row, UpsertMultiple (alt key) 34.6 ms/row. **Picker: not available.** In `$metadata`, `CreateMultiple` / `UpdateMultiple` / `UpsertMultiple` / `DeleteMultiple` are bound to `Collection(mscrm.crmbaseentity)` (the entity **set**). Nothing is bound to `pa_orderimport` itself. The connector's *Perform a bound action* (`PerformBoundAction`) always calls `{entitySet}({recordId})/{actionName}`, a **single record**, so collection-bound actions can't appear in or be called through it. The only other raw operation in the connector is *Perform a changeset request* (`ExecuteChangeset`, `$batch`). → For the demo flows, call CreateMultiple/UpsertMultiple with an HTTP action authenticated to Dataverse, e.g. *HTTP with Microsoft Entra ID (preauthorized)*; this hasn't been tested yet. |
| 8 | SharePoint `$batch` limit (100/500/1000/2000) + changeset atomicity | ⚠️ | List **PocImport** (Number `Amount`) on https://davidlorenzolopez.sharepoint.com. **100**: 100/100 ×4 runs (9.2–14.1 s). **500**: 500/500 ×2 (55.8 / 60.5 s). **1000**: **partial** ×2: 941/1000 (95.6 s) and 873/1000 (93.0 s). The outer `$batch` returns HTTP 200; SharePoint processes the parts **in order**, and every part after ~90 s of server time returns `500 Microsoft.SharePoint.Client.ClientServiceTimeoutException: The operation has timed out.` The limit is **time (~90 s per `$batch`)**, not a fixed item count, at ~95–110 ms/item. → **Use ≤ 500 items per `$batch`**, and count 5xx parts as well as 4xx (the export only shows 4xx). 2000 not run (1000 already times out). **Changeset atomicity: ❌ not atomic.** 100 items in one `changeset_` with item #1 invalid → HTTP 200, **99 × 201 + 1 × `400 InvalidClientQueryException` (*Cannot convert a primitive value to the expected type 'Edm.Double'*)**. The bad part was the **first** in the changeset and the remaining 99 were still committed, so SharePoint REST `$batch` has no rollback (unlike Dataverse, script 03). → Demo flows must treat a SharePoint batch as *best effort*: parse every part's status (2xx / 4xx / 5xx) and retry or report the failed ones. |
| 9 | flowrun cross-check + duration unit | ✅ | **`flowrun.duration` is in milliseconds** (T4: flow 60677 vs flowrun 61248; T5a: 150983 vs 151358; T3: 1862 vs 2378). flowrun is ~0.4–0.6 s longer, because it includes the trigger. New runs appear in `flowrun` only after a delay of a few minutes (some still missing 5+ min later). `flowrun.status` reflects the **whole run**: the T5b runs are *Failed* there while the flow's own telemetry says Succeeded (see #5), and T8-1000 *Partial* runs are *Failed*. `workflow()` tags include `xrmWorkflowId`, `logicAppName` (flow id), `environmentName`, `billingContext`, `licenseContextConfigName` (Premium). |

### Issues hit during setup

1. **`01-setup-dataverse.mjs`: 404 right after creating a table.** `GET EntityDefinitions(LogicalName='pa_benchmarkrun')/Attributes` returned *Entity … does not exist* immediately after the create succeeded, and the same happened for `pa_orderimport`. The entity GET can succeed before `/Attributes` does. **Fixed:** the script now polls `/Attributes` for up to ~60 s after create.
2. **`01-setup-dataverse.mjs`: one transient 400** (*An unexpected error occurred*) adding `pa_amount` (Decimal) right after the previous column. The identical payload worked when retried. Not reproduced, so not changed. A retry on 400/"unexpected error" for metadata POSTs may be worth adding.
3. **`pa app init` did not write a solution to `power.config.json`.** The app was pushed with `pa app push --solution-id <FlowOptimizationLab solutionid>` and landed in the solution.
4. **`pa` on headless Linux needs a Secret Service keyring.** `pa auth login` fails with `CachePersistenceError … GnomeKeyringError` because `usePlaintextFileOnLinux` is hard-coded to false. Workaround: run `dbus-daemon` + `gnome-keyring-daemon --unlock`, then `pa auth login --device-code`.
5. **The bound SharePoint connection was broken.** The existing connection (created 2025-07) had status *Error / Unauthorized*: `AADSTS50173 grant expired/revoked`. Binding it to the connection reference succeeded anyway, so a bound reference does not mean a healthy connection. The user fixed it in the portal. Worth a preflight check (`statuses[].status == Connected`) in `00-preflight` / `02-create-poc-flows`.
6. **Script 03, first run: `ECONNRESET`** at step 7/8, caused by a network proxy tunnel recycle in the container, not Dataverse. Rerun succeeded. The interrupted run left 799 rows, which were deleted afterwards with parallel single DELETEs (8 at a time, 0 failures, 0 remaining). The script only prints results at the very end and only cleans up on success, so writing partial results and cleaning up in a `finally` would make it more robust.

## App export

Published app (Local Play not tested; see #6).

### PoC results (2026-11-14T18:33:12.975Z)

Environment: df93e55c-cf3b-e261-87e1-ddb2b052de1d · App URL: https://apps.powerapps.com/play/e/df93e55c-cf3b-e261-87e1-ddb2b052de1d/app/fb0ed9e0-6b0e-43ab-8ff8-2a7c4687fdc6

| Test | Checklist | Verdict | Notes | Correlation |
|---|---|---|---|---|
| T3 Invoke a flow from the Code App | #3 | pass | Run() ok in 6.9 s; response: {"runid":"08584069262310378138969222704CU12","runurl":"https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/88346a50-54fa-4bba-ad73-2e885060481d/runs/08584069262310378138969222704CU12","respondedat":"2026-11-14T18:24:16.7972880Z"} | e401d214-09ca-4e7d-83fb-9d0f362b8572 |
| T4 Respond early, keep working (60 s) | #4 | pass | Run() resolved after 1.2 s (early ✅); flow duration 60.7 s (kept running ✅) | 80427677-b90a-4177-b3fc-36e5379bae77 |
| T5a > 120 s, no Respond action (150 s) | #5 | pass | Run() ok after 0.7 s; flow Succeeded in 151.0 s | c08a8058-9baf-40de-b5ce-33be2aa33508 |
| T5b > 120 s, Respond at the end (150 s) | #5 | info | Run() error: {"code":"UnexpectedError","message":"An unexpected error occurred."} after 120.3 s; flow Succeeded in 150.4 s | 2dfdde61-e044-4e37-9bc5-368625e39165 |
| T8-100 SharePoint $batch (100 items) | #8 | pass | HTTP 200, created 100/100, 4xx 0, batch call 14062 ms, payload 38855 bytes | 19c653b5-6f42-4a7d-ae9b-df09281451c6 |

#### Links (#6)

- Run (flows path): not tested · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069257374284637784635179CU05`
- Run (solutions/~preferred path): not tested · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/solutions/~preferred/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069257374284637784635179CU05`
- Edit flow (/edit): not tested · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/edit`
- Flow details (/details): not tested · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/details`
- Run (built in app from getContext): not tested · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069257374284637784635179CU05`

#### Generated services

- table: Pa_benchmarkrunsService (../generated/services/Pa_benchmarkrunsService.ts)
- flow: PocNoRespondService (../generated/services/PocNoRespondService.ts)
- flow: PocRespondEarlyService (../generated/services/PocRespondEarlyService.ts)
- flow: PocRespondLateService (../generated/services/PocRespondLateService.ts)
- flow: PocSpBatchService (../generated/services/PocSpBatchService.ts)

### PoC results (2026-11-14T18:55:54.516Z → 19:05:01.744Z), second session (final export adds T8-100-bad; links re-checked with the same verdicts)

| Test | Checklist | Verdict | Notes | Correlation |
|---|---|---|---|---|
| T3 Invoke a flow from the Code App | #3 | pass | Run() ok in 2.1 s; response: {"runid":"08584069245352739383743956705CU31","runurl":"https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/88346a50-54fa-4bba-ad73-2e885060481d/runs/08584069245352739383743956705CU31","respondedat":"2026-11-14T18:52:31.8728068Z"} | ab333e12-fab9-45cd-8cfd-ec7401179317 |
| T4 Respond early, keep working (60 s) | #4 | pass | Run() resolved after 0.8 s (early ✅); flow duration 60.8 s (kept running ✅) | 61ed0e8a-a6e0-463a-8bbb-b437b1dd7dbf |
| T5a > 120 s, no Respond action (150 s) | #5 | pass | Run() ok after 0.5 s; flow Succeeded in 150.2 s | e9f65758-4260-40e3-87f9-ac47fd65c7a2 |
| T5b > 120 s, Respond at the end (150 s) | #5 | info | Run() error: {"error":{"code":"TimedOut","message":"Invocation of API timed out.","details":[]}} after 120.1 s; flow Succeeded in 150.5 s | 78b95fda-c309-484d-a46d-3d6564c6db52 |
| T8-100 SharePoint $batch (100 items) | #8 | pass | HTTP 200, created 100/100, 4xx 0, batch call 9173 ms, payload 38855 bytes | c187ad84-97f6-460d-8f5a-002d1eebdb4d |
| T8-500 SharePoint $batch (500 items) | #8 | pass | HTTP 200, created 500/500, 4xx 0, batch call 55804 ms, payload 194055 bytes | 2bb246ae-71a9-4deb-a6c1-edfd4e6b9a1a |
| T8-1000 SharePoint $batch (1000 items) | #8 | info | HTTP 200, created 873/1000, 4xx 0, batch call 92997 ms, payload 388055 bytes | f08e88f8-d33e-4dce-a38a-cee3e1c3a923 |
| T8-100-bad SharePoint $batch (100 items, 1 invalid) | #8 | info | HTTP 200, created 99/100, 4xx 1, batch call 9529 ms, payload 38868 bytes | 2d49df1e-6d3f-4b5d-ad61-23baebe476a5 |

#### Links (#6)

- Run (flows path): ok · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069244388952803805661278CU05`
- Run (solutions/~preferred path): broken · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/solutions/~preferred/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069244388952803805661278CU05`
- Edit flow (/edit): broken · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/edit`
- Flow details (/details): ok · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/details`
- Run (built in app from getContext): ok · `https://make.powerautomate.com/environments/df93e55c-cf3b-e261-87e1-ddb2b052de1d/flows/7c753364-30d6-1edc-f62f-707e0ec09d00/runs/08584069244388952803805661278CU05`

T8-1000 detail (from the `pa_benchmarkrun` row and the run's `Response_text` output): run `08584069244388952803805661278CU05` got 873 × `201 Created`, then parts 874–1000 × `500 Internal Server Error`, `{"error":{"code":"-2146233083, Microsoft.SharePoint.Client.ClientServiceTimeoutException","message":{"lang":"en-US","value":"The operation has timed out."}}}`. The earlier 1000 run `08584069245901406068699004510CU31` got 941 × 201, 59 × 5xx. `pa_error` only stores the first 4,000 characters of the batch response, so the 5xx parts are only visible in the run history.

T8-100-bad detail: run `08584069239050508413196862487CU30` (`pa_status` = Partial): part 1 → `400 Bad Request`, `{"error":{"code":"-1, Microsoft.SharePoint.Client.InvalidClientQueryException","message":{"lang":"en-US","value":"Cannot convert a primitive value to the expected type 'Edm.Double'. ..."}}}`; parts 2–100 → `201 Created`.

## Script 03 output

`node poc/scripts/03-bulk-api-test.mjs 1000`. Client-side timings from the cloud container (US) to PREMIUM-US. Raw: `poc/results/bulk-api-poc-2026-11-14T18-15-00-075Z.json`.

| test | records | HTTP requests | ms | ms/record | note |
|---|---:|---:|---:|---:|---|
| Single POST (sequential) | 100 | 100 | 18212 | 182.1 | extrapolated to N=1000: ~182 s |
| CreateMultiple (1000/request) | 1000 | 1 | 10141 | 10.1 | |
| UpdateMultiple (1000/request) | 1100 | 2 | 11788 | 10.7 | |
| UpsertMultiple (alt key, 50% new) | 1000 | 1 | 34551 | 34.6 | |
| `$batch` changeset 100 creates (valid) | 100 | 1 | 2434 | 24.3 | HTTP 200, rows created: 100 |
| `$batch` changeset 100 creates (1 invalid) | 100 | 1 | 946 | 9.5 | HTTP 400, rows created: 0 → ✅ rolled back |
| `$batch` no changeset + continue-on-error (1 invalid) | 100 | 1 | 2951 | 29.5 | HTTP 200, rows created: 99 (expect 99) |
| `$batch` DELETE (1000/request) | 1799 | 2 | 516853 | 287.3 | remaining: 0 |

Observations:
- CreateMultiple is ~18× faster per record than sequential single creates. UpsertMultiple through the alternate key is ~3.4× slower than CreateMultiple.
- Changeset atomicity works as expected: 1 invalid record rolls back the whole changeset (0 rows). Without a changeset, continue-on-error gives 99/100.
- `$batch` DELETE is very slow (~287 ms/record, 8.6 min for 1799 rows), slower than sequential single POSTs. For cleanup, prefer `BulkDelete` or parallel deletes.

## Script 04 output

`node poc/scripts/04-check-runs.mjs 10`, run right after the app tests (the two newest runs weren't in `flowrun` yet):

```text
workflow() tags exposed by the platform: {
  flowDisplayName: 'PocRespondLate',
  billingContext: 'DynamicsSolution',
  capabilities: 'Premium,DVConnector',
  licenseContextConfigName: 'CONTEXT_USER-241D68385AFA43199C16A00EB6A78D16_POWERAPPS_PREMIUM40K',
  environmentName: 'df93e55c-cf3b-e261-87e1-ddb2b052de1d',
  logicAppName: '88346a50-54fa-4bba-ad73-2e885060481d',
  environmentWorkflowId: 'df93e55c:2Dcf3b:2De261:2D87e1:2Dddb2b052de1d-162174b3:2D56e0:2Df111:2D84e8:2D7ced8d6f7afa',
  xrmWorkflowId: '162174b3-56e0-f111-84e8-7ced8d6f7afa',
  environmentFlowSuspensionReason: 'df93e55c:2Dcf3b:2De261:2D87e1:2Dddb2b052de1d-None',
  sharingType: 'CommonDataService',
  state: 'Enabled',
  createdTime: '11/14/2026 6:10:58 PM',
  lastModifiedTime: '11/14/2026 6:10:58 PM',
  createdBy: '241d6838-5afa-4319-9c16-a00eb6a78d16',
  triggerType: 'Instant',
  createdByUserId: 'C7257B319BE54652957D9E2DFDC851E2-DF93E55CCF3BE26187E1DDB2B052DE1D-ENV'
} 

┌─────────┬────────────────────────┬────────────────┬─────────────┬────────┬─────────────────┬───────────────┬─────────────────────────────────────┐
│ (index) │ created                │ variant        │ status      │ flowMs │ flowrunDuration │ flowrunStatus │ runId                               │
├─────────┼────────────────────────┼────────────────┼─────────────┼────────┼─────────────────┼───────────────┼─────────────────────────────────────┤
│ 0       │ '2026-11-14T18:32:28Z' │ 'SpBatch'      │ 'Succeeded' │ 15026  │ '(not found)'   │ ''            │ '08584069257374284637784635179CU05' │
│ 1       │ '2026-11-14T18:29:48Z' │ 'RespondLate'  │ 'Succeeded' │ 150444 │ '(not found)'   │ ''            │ '08584069258975033739693487173CU31' │
│ 2       │ '2026-11-14T18:27:10Z' │ 'NoRespond'    │ 'Succeeded' │ 150983 │ 151358          │ 'Succeeded'   │ '08584069260549339911896690580CU24' │
│ 3       │ '2026-11-14T18:24:58Z' │ 'RespondEarly' │ 'Succeeded' │ 60677  │ 61248           │ 'Succeeded'   │ '08584069261876531666149141405CU02' │
│ 4       │ '2026-11-14T18:24:14Z' │ 'RespondLate'  │ 'Succeeded' │ 1862   │ 2378            │ 'Succeeded'   │ '08584069262310378138969222704CU12' │
└─────────┴────────────────────────┴────────────────┴─────────────┴────────┴─────────────────┴───────────────┴─────────────────────────────────────┘

ℹ️  Compare flowMs with flowrunDuration to confirm the unit of flowrun.duration (docs disagree: seconds vs ms).
```

Second run, `node poc/scripts/04-check-runs.mjs 20`, after the second app session. `status` = the flow's own `pa_benchmarkrun` row; `flowrunStatus` = platform:

```text
┌─────────┬────────────────────────┬────────────────┬─────────────┬────────┬─────────────────┬───────────────┬─────────────────────────────────────┐
│ (index) │ created                │ variant        │ status      │ flowMs │ flowrunDuration │ flowrunStatus │ runId                               │
├─────────┼────────────────────────┼────────────────┼─────────────┼────────┼─────────────────┼───────────────┼─────────────────────────────────────┤
│ 0       │ '2026-11-14T18:54:06Z' │ 'SpBatch'      │ 'Partial'   │ 93877  │ '(not found)'   │ ''            │ '08584069244388952803805661278CU05' │
│ 1       │ '2026-11-14T18:53:03Z' │ 'SpBatch'      │ 'Succeeded' │ 56632  │ 56949           │ 'Succeeded'   │ '08584069245025556166037268363CU13' │
│ 2       │ '2026-11-14T18:52:36Z' │ 'SpBatch'      │ 'Succeeded' │ 9825   │ 10100           │ 'Succeeded'   │ '08584069245294190433141747002CU24' │
│ 3       │ '2026-11-14T18:52:33Z' │ 'RespondLate'  │ 'Succeeded' │ 150458 │ '(not found)'   │ ''            │ '08584069245319368439059362230CU02' │
│ 4       │ '2026-11-14T18:52:32Z' │ 'NoRespond'    │ 'Succeeded' │ 150235 │ '(not found)'   │ ''            │ '08584069245335927589037623775CU29' │
│ 5       │ '2026-11-14T18:52:31Z' │ 'RespondEarly' │ 'Succeeded' │ 60849  │ 61117           │ 'Succeeded'   │ '08584069245344481776327621311CU20' │
│ 6       │ '2026-11-14T18:52:30Z' │ 'RespondLate'  │ 'Succeeded' │ 1460   │ '(not found)'   │ ''            │ '08584069245352739383743956705CU31' │
│ 7       │ '2026-11-14T18:51:35Z' │ 'SpBatch'      │ 'Partial'   │ 96498  │ 96818           │ 'Failed'      │ '08584069245901406068699004510CU31' │
│ 8       │ '2026-11-14T18:51:32Z' │ 'RespondLate'  │ 'Succeeded' │ 1238   │ 4580            │ 'Succeeded'   │ '08584069245935228637616034665CU02' │
│ 9       │ '2026-11-14T18:49:53Z' │ 'SpBatch'      │ 'Succeeded' │ 61253  │ 61552           │ 'Succeeded'   │ '08584069246924149662753557928CU04' │
│ 10      │ '2026-11-14T18:49:38Z' │ 'NoRespond'    │ 'Succeeded' │ 150253 │ '(not found)'   │ ''            │ '08584069247067924379361118640CU10' │
│ 11      │ '2026-11-14T18:49:38Z' │ 'RespondLate'  │ 'Succeeded' │ 150498 │ 150931          │ 'Failed'      │ '08584069247069561426864487478CU17' │
│ 12      │ '2026-11-14T18:49:37Z' │ 'RespondEarly' │ 'Succeeded' │ 60458  │ 60817           │ 'Succeeded'   │ '08584069247082390103994655809CU20' │
│ 13      │ '2026-11-14T18:49:36Z' │ 'RespondLate'  │ 'Succeeded' │ 1219   │ 1719            │ 'Succeeded'   │ '08584069247090372644552350978CU26' │
│ 14      │ '2026-11-14T18:49:33Z' │ 'SpBatch'      │ 'Succeeded' │ 13647  │ 13880           │ 'Succeeded'   │ '08584069247121551634188087875CU05' │
│ 15      │ '2026-11-14T18:48:38Z' │ 'SpBatch'      │ 'Succeeded' │ 11413  │ 11889           │ 'Succeeded'   │ '08584069247675234871400894509CU10' │
│ 16      │ '2026-11-14T18:45:32Z' │ 'RespondLate'  │ 'Succeeded' │ 150284 │ 151060          │ 'Failed'      │ '08584069249528720793921157646CU28' │
│ 17      │ '2026-11-14T18:42:47Z' │ 'NoRespond'    │ 'Succeeded' │ 150408 │ 150706          │ 'Succeeded'   │ '08584069251178251465952790015CU08' │
│ 18      │ '2026-11-14T18:41:39Z' │ 'RespondEarly' │ 'Succeeded' │ 60766  │ 61089           │ 'Succeeded'   │ '08584069251865658327424668595CU05' │
│ 19      │ '2026-11-14T18:38:23Z' │ 'RespondLate'  │ 'Succeeded' │ 1581   │ 2041            │ 'Succeeded'   │ '08584069253821045991061283629CU15' │
└─────────┴────────────────────────┴────────────────┴─────────────┴────────┴─────────────────┴───────────────┴─────────────────────────────────────┘

ℹ️  Compare flowMs with flowrunDuration to confirm the unit of flowrun.duration (docs disagree: seconds vs ms).
```
