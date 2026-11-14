# Demo Scenarios: Slow vs Fast Flow Pairs

Every use case ships as **at least two flows** with the same input and the same expected output:

- `UCx-SLOW`: the anti-pattern, the way we usually find it in customer tenants
- `UCx-FAST`: the optimized rebuild
- (optional) `UCx-MID`: an intermediate step, e.g. "same loop but concurrency 50", to show that tuning helps but redesign wins

Every pair must produce an **identical result checksum** (e.g. total amount, record count, hash of IDs), so the audience sees *same result, different cost*.

## Demo dataset ("realistic enterprise volumes")

| Store | Entity | Volume | Notes |
|---|---|---|---|
| SharePoint list `Orders` | OrderId, Customer (lookup), Region, Status, Amount, OrderDate, ~20 extra "noise" columns | 20,000 items | Index Status, Region, OrderDate. The wide columns make `$select` visibly matter |
| SharePoint list `Customers` | CustomerId, Name, Segment, Country | 2,000 items | Lookup target |
| SharePoint list `OrderImport` | target of bulk create/delete | 0 → 1,000 / 5,000 | Reset between runs |
| Dataverse table `pa_order` | same shape | 20,000 rows | |
| Dataverse table `pa_orderimport` | bulk target | 0 → 5,000 | Standard table (CreateMultiple/UpdateMultiple/UpsertMultiple) |
| Dataverse elastic table `pa_telemetry` (optional) | bulk target | | Shows DeleteMultiple (elastic only) |
| Excel table (optional) | `Orders.xlsx` | 2,000 rows | For the Office Scripts bonus |

Seed and reset flows: `SEED-SP`, `SEED-DV`, `RESET-<UC>`. These are also launched from the Code App.

**Volume parameter.** Every flow takes `volume` as an input. Slow flows at full volume can take minutes or hours, so the app lets us run *SLOW @ 500* live and show *SLOW @ 5,000* from a pre-recorded run in the history (the harness keeps all runs). It extrapolates the cost linearly and labels the figure as an extrapolation.

---

## Use cases

| UC | Theme (abstract section) | SLOW flow | FAST flow | Main KPI to show |
|---|---|---|---|---|
| **UC1** | Data shaping | Apply to each over 5,000 orders → Condition (Status = Open) → Compose new object → **Append to array variable** | **Filter array** + **Select** (advanced mode, `addProperty`) | duration, action count (≈ 3×N vs 2) |
| **UC2** | Join / enrichment ➕ | For each order → **Get item** (customer) or nested Apply to each over customers | **Dictionary pattern**: Select customers → `json(concat('{',join(...),'}'))` → Select orders with `outputs('Dict')?[CustomerId]` | duration, API calls (N vs 2) |
| **UC3** | XPath aggregation | Apply to each → **Increment variable** for count, Set variable for sum/min/max, group by region with nested loops | Select → `xml()` → **xpath** sum/count/avg + `max()/min()`; group-by via `union()` + Select | duration, actions, identical totals |
| **UC4** | SharePoint bulk | Apply to each → **Create item** (concurrency 1) for 1,000 items; MID: concurrency 50 (show throttling 429s) | **`$batch`** via Send HTTP request to SharePoint, `chunk()` + Select-built changesets | duration, HTTP requests (1,000 vs 1–10), 429 retries |
| **UC5** | Dataverse bulk | Apply to each → **Add a new row** for 5,000 rows; update and delete the same way | **CreateMultiple / UpdateMultiple / UpsertMultiple** (1,000 per call); `$batch` changeset for parent+children all-or-nothing; delete: `$batch` DELETE vs BulkDelete | duration, requests, transactional demo (inject 1 bad row → full rollback) |
| **UC6** | Filter at source | Get items (pagination 100k, all columns) → Filter array / Condition → Get item per row for the customer lookup | **Filter Query** + **$select** + **$expand** (lookup) + `$top`; DV: FetchXML aggregate / `$apply` | duration, **payload KB**, items retrieved vs used |
| **UC7** | Trigger conditions | "When an item is modified" → Condition → Terminate if not relevant (runs every time) | Same flow with a **trigger condition**: runs only when Status changes to Approved | **number of runs** and billed actions for 100 modifications (only ~10 relevant) |
| **UC8** | Platform tuning | (a) Concurrent loop + **Set variable** → wrong results (race); (b) Get items without pagination → silently 100 items; (c) variables everywhere | (a) Compose / Select inside the loop; (b) correct pagination threshold; (c) Compose | correctness flags, duration |
| **UC9** ➕ bonus | Excel | Apply to each → **Add a row into a table** (2,000 rows) | **Office Script** `setValues()` in one call | the most dramatic gap; good opener or closer |
| **UC10** ➕ bonus | Fan-out | One flow loops 20,000 records sequentially | `chunk()` → Apply to each (concurrency 10) → **child flow** per chunk | wall-clock time |

Recommended live set for 50 minutes: **UC1, UC3, UC4 (or UC5), UC6, UC7**, with UC8 sprinkled into the tuning section and UC9 as the opener ("hours to seconds" hook). The rest are backups or Q&A material.

---

## Draft agenda (50 min)

| Min | Block | Demo |
|---|---|---|
| 0–4 | Hook: "this flow took 5 hours". Why slow flows cost money (PPR limits, throttling, trust) | UC9 or UC1 launched from the app, slow and fast side by side |
| 4–8 | How to find the bottleneck: run history, action timings, flowrun table, Automation Center | Open the SLOW run via the app link |
| 8–17 | Kill loops: Select / Filter array / dictionary join + decision framework | UC1, UC2 |
| 17–23 | XPath aggregation | UC3 |
| 23–33 | Bulk: SharePoint `$batch`, Dataverse bulk messages + changesets | UC4, UC5 (rollback demo) |
| 33–41 | Filter at source + trigger conditions | UC6, UC7 |
| 41–46 | Tune the platform: concurrency (and the race-condition trap), Compose vs variables, pagination | UC8 |
| 46–50 | Scoreboard in the app (all runs, speedup factors), takeaways, Q&A | Code App dashboard |

---

## What each flow must record (contract with the Code App)

See `03-code-app-benchmark-harness.md`. Each flow writes start/end timestamps, records processed, loop iterations, HTTP/connector calls, an estimated action count, the result checksum and the run links into the `pa_benchmarkrun` Dataverse table.
