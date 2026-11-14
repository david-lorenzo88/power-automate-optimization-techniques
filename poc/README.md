# PoC: Code App benchmark harness checklist

This kit runs checklist items 1–9 from [`docs/03-code-app-benchmark-harness.md`](../docs/03-code-app-benchmark-harness.md) against **your** tenant.
It takes about 45 minutes, and most of that is waiting for the 60 s / 150 s test flows.

| Piece | What it does |
|---|---|
| `scripts/00-preflight.mjs` | Dataverse access, WhoAmI, flowrun table readable, publisher prefix free |
| `scripts/01-setup-dataverse.mjs` | Publisher `flowoptlab` (prefix `pa`), solution **FlowOptimizationLab**, tables `pa_benchmarkrun` (telemetry) + `pa_orderimport` (bulk tests, alt key), connection references for Dataverse + SharePoint |
| `scripts/02-create-poc-flows.mjs` | Creates the 4 PoC flows **in the solution** via the Web API (`flows/definitions.mjs`) and turns them on |
| `scripts/03-bulk-api-test.mjs` | Dataverse single vs CreateMultiple / UpdateMultiple / UpsertMultiple / `$batch`, plus a changeset rollback proof (client-side timings) |
| `scripts/04-check-runs.mjs` | Compares flow-reported durations with the platform `flowrun` table and prints the `workflow()` tags |
| `../app` | Code App "Flow Optimization Lab · PoC": runs T3, T4, T5a, T5b, T8, tests the links, exports results as Markdown |

PoC flows (they all take a **single text input `payload`**, i.e. a JSON string):

| Flow | Shape | Used by |
|---|---|---|
| `PocRespondEarly` | create row → **Respond** → Delay → update row | T4 |
| `PocRespondLate` | create row → Delay → update row → **Respond** | T3 (delay 1 s), T5b (150 s) |
| `PocNoRespond` | create row → Delay → update row (no Respond) | T5a |
| `PocSpBatch` | build N changeset parts with **Select + join** (no loop) → one `_api/$batch` call → count `HTTP/1.1 201` | T8 |

---

## 0. Prerequisites (on your machine)

- Node.js LTS, Git, Azure CLI (`az login` with the presenter account)
- `npm install --global @microsoft/power-apps-cli` (provides `pa`, v1.4.1 at the time of writing)
- Power Platform admin center → your environment → Settings → Product → Features → **Enable code apps = On** (checklist #1)
- Presenter account has Power Apps Premium (checklist #1)
- A SharePoint site where you can create a list (for T8)

```bash
export DATAVERSE_URL=https://<yourorg>.crm<N>.dynamics.com
```

PowerShell: `$env:DATAVERSE_URL="https://<yourorg>.crm<N>.dynamics.com"`

## 1. Preflight + Dataverse setup (#1, #2)

```bash
node poc/scripts/00-preflight.mjs
```

```bash
node poc/scripts/01-setup-dataverse.mjs
```

Then go to make.powerapps.com → Solutions → **Flow Optimization Lab** → Connection references. Open `FlowOptLab Dataverse` and `FlowOptLab SharePoint` and select (or create) a connection for each one.

## 2. Create the PoC flows (#2)

```bash
node poc/scripts/02-create-poc-flows.mjs
```

The script prints the **flow id** (`resourceid`) of each flow. If a flow can't be turned on, open it from the solution, fix whatever the designer flags (usually the connection), save it and turn it on. Then re-run the script; it's idempotent.

> If the Web API creation doesn't work in your tenant, build the flows by hand from `poc/flows/generated/*.json` (each action's inputs are in there) and **tell me what failed**. I'll fix the generator, since we'll use the same approach for the ~20 demo flows.

## 3. SharePoint list for T8 (#8)

On your site, create a list **PocImport** and add a **Number** column called **Amount**.

## 4. Code App (#3–#6, #8)

```bash
cd app
npm install
pa auth login
pa app init --display-name "Flow Optimization Lab" --environment-id <environment-id>
pa app add data-source --connector dataverse --table pa_benchmarkrun
pa app list-flows --search Poc
pa app add flow --flow-id <PocRespondEarly id>
pa app add flow --flow-id <PocRespondLate id>
pa app add flow --flow-id <PocNoRespond id>
pa app add flow --flow-id <PocSpBatch id>
npm run dev        # open the "Local Play" URL it prints
```

Before publishing, check that `pa app init` added your solution to `power.config.json`. If it didn't, pass it explicitly when you push. Then:

```bash
npm run build
pa app push
```

In the app:

1. **#1 Context** should show your environment, user and Dataverse URL.
2. **#3 Generated services** should list 1 table + 4 flows. The app finds them by name, so whatever file names the generator picks don't matter.
3. Run **T3**, **T4**, **T5a**, **T5b** (you can run them in parallel; T5 takes about 2.5 min).
4. **#6 Links**: click each link both ways (anchor and window.open) and mark it *works* / *broken*. Do this **once in Local Play and once in the published app**, because the CSP may differ.
5. **T8**: enter the site URL and run batch sizes 100 → 500 → 1000 → 2000, then 100 with *inject 1 invalid item* (the changeset atomicity question).
6. **Export** → *Copy results as Markdown*.

## 5. Bulk API + platform cross-check (#7, #9)

```bash
node poc/scripts/03-bulk-api-test.mjs 1000
```

```bash
node poc/scripts/04-check-runs.mjs 10
```

For #7 also check by hand: in any flow, add the Dataverse action **Perform a bound action**, set Table = *Order imports*, and look in the Action Name list for **CreateMultiple / UpdateMultiple / UpsertMultiple**.

## 6. Send me the results

Paste these into `poc/RESULTS.md` and commit it, or just paste them in the chat:
- the app's Markdown export (published app + local)
- the output of scripts 03 and 04 (03 also saves `poc/results/*.json`)
- the #7 picker finding
- any error messages

I'll update the design (doc 03) and generate the ~20 demo flows from the parts that worked.
