# Flow Optimization Lab: slow vs fast flow pairs

The demo flows for [`docs/02-demo-scenarios.md`](../docs/02-demo-scenarios.md), built on the harness design in
[`docs/03-code-app-benchmark-harness.md`](../docs/03-code-app-benchmark-harness.md) and the PoC findings in
[`poc/RESULTS.md`](../poc/RESULTS.md). Every use case has a SLOW flow (the anti-pattern), a FAST flow (the rebuild),
sometimes a MID flow (tuned but not redesigned), the same input and the **same result checksum**.

Everything is generated: flow definitions are code (`lab/flows/*.mjs`), and the scripts deploy them into the
**Flow Optimization Lab** solution through the Dataverse Web API.

| Piece | What it does |
|---|---|
| `scripts/10-setup-lab.mjs` | SharePoint lists `Customers`, `Orders` (26 columns, indexed), `OrderImport`, `UC7Items`; Dataverse table `pa_order`; connection references `pa_httpdataverse` + `pa_excel`. Writes `lab.config.json` |
| `scripts/11-upload-uc9-assets.mjs` | UC9: uploads two copies of `assets/Uc9Orders.xlsx` and the Office Scripts in `assets/office-scripts/` (wrapped as `.osts`) to `Shared Documents/FlowOptLab` |
| `scripts/20-deploy-flows.mjs` | Creates/updates the 30 flows in the solution and turns them on. Writes `flows/deployed.json` and `flows/generated/*.json` |
| `scripts/30-run.mjs` | Runs any lab flow from the command line (through `LAB-TestRunner`) and prints its KPIs; `--pair UCx` races all variants and compares |
| `flows/lib.mjs` | The telemetry contract every flow shares, connector helpers, deterministic demo data formulas |
| `flows/scenarios.mjs` | The flow definitions (UC1–UC10, seed, reset, test runner) |
| `../app` | Code App: **Scenarios** tab (run SLOW/MID/FAST, live KPIs, speed-up, checksum match, scoreboard), **Seed & reset**, **PoC checks** |

## Use cases

| UC | SLOW | MID | FAST | Data |
|---|---|---|---|---|
| UC1 Data shaping | Apply to each → Condition → Compose → Append to array | | Filter array + Select | SP `Orders` 1…N |
| UC2 Join | Apply to each → query the customer per order | | Dictionary: `json(concat('{', join(Select), '}'))` + key lookup | SP `Orders`, `Customers` |
| UC3 Aggregation | Loop per region × loop over orders, Increment/Set variable | | `xml()` + `xpath` count/sum, `min()`/`max()`, `union()` group-by | SP `Orders` |
| UC4 SharePoint bulk | Create item, concurrency 1 | concurrency 50 | `$batch` (250 per changeset, 4 in parallel) via Send HTTP request | SP `OrderImport` |
| UC5 Dataverse bulk | Add a new row, concurrency 1 | concurrency 50 | `CreateMultiple`, 1,000 per request (HTTP with Entra ID) | DV `pa_orderimport` |
| UC6 Filter at source | Get items (all columns, paginated) → Filter array → Get item per row | | REST `$filter` + `$select` + `$expand` | SP `Orders` |
| UC6DV Aggregate at source | List rows (all columns) → Filter array → sum | | FetchXML aggregate | DV `pa_order` |
| UC7 Trigger conditions | automated: runs on every change, Condition + Terminate | | automated: trigger conditions | SP `UC7Items`, driven by `UC7-DRIVER` |
| UC8 Tuning traps | Get items **without** pagination (100 rows) + concurrent loop on a variable | | pagination + Select/xpath | SP `Orders` |
| UC9 Excel | Add a row into a table per row | | Office Script `setValues()` | `Uc9Slow.xlsx` / `Uc9Fast.xlsx` |
| UC10 Fan-out | one sequential loop | | `chunk()` → concurrency 10 → child flow `UC10-CHILD` | in memory |

Admin flows: `SEED-SP` (Customers 2,000, Orders 20,000, UC7Items 100 via `$batch`), `SEED-DV` (`pa_order` 20,000 via
CreateMultiple), `RESET` (`payload.target` = `uc4` | `uc5` | `uc7` | `uc9` | `all`).

### What every flow records

Each app-launched flow follows one contract (`flows/lib.mjs`):

1. Power Apps V2 trigger with a single text input `payload` (JSON: `correlationId`, `volume`, `label`, options).
2. `Create_run_row` (status *Running*) in `pa_benchmarkrun`, then **respond early** (PoC #4/#5: `Run()` fails after 120 s).
3. The work runs in a `Work` scope and ends with a `Kpi` compose: records processed/retrieved, loop iterations,
   connector calls, estimated actions, payload bytes, **checksum**, and optionally its own status (*Partial*).
4. `Update_run_row` writes the KPIs, duration and errors. `Update_run_row_fallback` still closes the row if `Kpi` never ran.

Links are `/flows/{id}/runs/{runId}` and `/flows/{id}/details` (the `/solutions/~preferred/` and `/edit` forms were broken in the PoC).
`pa_editurl` holds the details URL.

## Setup

Prerequisites: the PoC setup (`poc/scripts/01-setup-dataverse.mjs`, connection references bound), `az login`,
`export DATAVERSE_URL=https://<org>.crm.dynamics.com`. Node ≥ 22.

```bash
node lab/scripts/10-setup-lab.mjs
```

```bash
node lab/scripts/11-upload-uc9-assets.mjs
```

Then in make.powerapps.com → Solutions → **Flow Optimization Lab** → Connection references, bind:

- `FlowOptLab HTTP Dataverse (Entra ID)` → new *HTTP with Microsoft Entra ID (preauthorized)* connection, Base Resource URL **and**
  Microsoft Entra ID Resource URI = your Dataverse URL (UC5-FAST, SEED-DV, RESET)
- `FlowOptLab Excel` → an *Excel Online (Business)* connection (UC9, RESET)

```bash
node lab/scripts/20-deploy-flows.mjs
```

```bash
node lab/scripts/30-run.mjs SEED-SP --timeout 3600
```

```bash
node lab/scripts/30-run.mjs SEED-DV --timeout 3600
```

Add the flows to the Code App (one `pa app add flow` per app-launched flow; ids are in `flows/deployed.json`), then build and push:

```bash
cd app && node -e "for (const [k,v] of Object.entries(require('../lab/flows/deployed.json'))) if (v.app) console.log(v.flowId)" | xargs -n1 pa app add flow --non-interactive --flow-id
```

```bash
npm run build && pa app push --solution-id <FlowOptimizationLab solutionid>
```

## Running from the command line

```bash
node lab/scripts/30-run.mjs --pair UC1 --volume 500
```

```bash
node lab/scripts/30-run.mjs UC5-FAST --volume 1000 --payload '{"injectBadItem":true}'
```

```bash
node lab/scripts/30-run.mjs RESET --payload '{"target":"uc4"}'
```

`30-run.mjs` calls `LAB-TestRunner` (HTTP trigger), which runs the flow as a child flow. The rows land in
`pa_benchmarkrun` exactly like app runs, so they show up in the app's scoreboard. The runner's trigger URL carries a
SAS signature: the script fetches it with your `az login` each time and never prints it.

## Platform behaviour found while building

| Finding | Consequence |
|---|---|
| A Power Apps V2 flow can't be started from the Flow management API (`triggers/manual/run` ignores the input), and the `shared_logicflows` connector only accepts the Power Apps player (connection ACL). | Test runner flow: HTTP trigger → *Run a Child Flow*. Power Apps V2 flows work as child flows. |
| Inside a SharePoint `$batch`, `X-HTTP-Method: MERGE` is ignored (*"The parameter __metadata does not exist in method GetById"*). | Updates in a batch use `PATCH` in the request line. |
| `result('<Apply to each>')` returns one entry **per action**, whose `outputs` is an array of every iteration. | UC10 collects child results with `first(filter(result(...)))?['outputs']`. |
| `add()` takes exactly two arguments. | Nested `add()` in KPI formulas. |
| A Scope whose last action succeeds is *Succeeded* even if an earlier action (a loop) failed. | Flows set their own status (`Partial`) when fewer records than requested were written. |
| A flow that was never turned on has no flow id (`resourceid`), so `pa app add flow` fails for it. | Bind connections → deploy → then `pa app add flow`. |
| Updating `clientdata` of a flow whose connection references aren't bound fails (400). | Deploy those flows after binding. |
| A SharePoint lookup column `Customer` is exposed in REST as `CustomerId`. | The text column is named `CustCode`. |
| `pa` on headless Linux needs a Secret Service keyring (see `poc/RESULTS.md`). | `dbus-daemon` + `gnome-keyring-daemon --unlock`, then `pa auth login --device-code`. |
| SharePoint processes a `$batch` serially (~100–170 ms per item) and stops after ~90 s of server time. | UC4-FAST sends batches of 250, 4 in parallel; SEED-SP does the same. |
| Dataverse `BulkDelete` with `RunNow: true` is refused (*does not support Synchronous Bulk Delete*). | RESET submits an asynchronous job (`RunNow: false`) with a typed `QueryExpression`. |
| An `If` whose condition is false reports *Succeeded*. | RESET reports each target as done / skipped / FAILED by looking at the action inside. |

## Results

See [`RESULTS.md`](RESULTS.md).
