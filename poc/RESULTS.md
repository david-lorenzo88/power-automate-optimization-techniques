# PoC results

Fill in after running the kit (see poc/README.md). ✅ / ❌ / ⚠️ + notes.

Run on 2026-11-14 against **PREMIUM-US** (`df93e55c-cf3b-e261-87e1-ddb2b052de1d`, https://premium-us.crm.dynamics.com, Sandbox), as david@powerplatformninja.com.
Scripts and `pa` ran from a Linux cloud container (Node 22.23, Azure CLI 2.92.0, `pa` 1.4.1), signed in with device code.

| # | Check | Result | Notes |
|---|---|---|---|
| 1 | Code apps enabled + Premium | ✅ | Confirmed by the user in PPAC. Preflight: WhoAmI OK (699 ms), `flowrun` readable, prefix `pa` free. |
| 2 | Solution, tables, flows created via script | ✅ | Publisher, solution, `pa_benchmarkrun`, `pa_orderimport` (+ alt key), 2 connection refs, and all 4 flows created **in the solution** via the Web API and turned on. Two setup errors, see *Issues*. |
| 3 | `pa app add flow` + `Run()` works (T3) | ⏳ | `pa app init`, `add data-source` and `add flow` ×4 all worked and generated `Poc*Service.ts`. `Run()` result pending (app run). |
| 4 | Respond early keeps the flow running (T4) | ⏳ | Run() resolve time: |
| 5 | > 120 s: no-respond (T5a) / respond-late (T5b) | ⏳ | Run() behaviour: |
| 6 | Links: flows path / solutions path / edit / details, anchor vs window.open, local vs published | ⏳ | Published app only. Local Play can't be tested from the cloud container, because the browser can't reach its `localhost:3000`. |
| 7 | CreateMultiple in "Perform a bound action" picker; script 03 timings | ⚠️ API ✅ / picker ⏳ | API: CreateMultiple 1000 rows in **10.1 s** in 1 request, vs ~182 s extrapolated for sequential single POSTs (~18×). UpdateMultiple 10.7 ms/row, UpsertMultiple (alt key) 34.6 ms/row. Picker check pending. |
| 8 | SharePoint `$batch` limit (100/500/1000/2000) + changeset atomicity | ⏳ | List **PocImport** (Number column `Amount`) created on https://davidlorenzolopez.sharepoint.com via the bound SharePoint connection. |
| 9 | flowrun cross-check + duration unit | ⏳ | Runs after the app tests. |

### Issues hit during setup

1. **`01-setup-dataverse.mjs`: 404 right after creating a table.** `GET EntityDefinitions(LogicalName='pa_benchmarkrun')/Attributes` returned *Entity … does not exist* immediately after the create succeeded, and the same happened for `pa_orderimport`. The entity GET can succeed before `/Attributes` does. **Fixed:** the script now polls `/Attributes` for up to ~60 s after create.
2. **`01-setup-dataverse.mjs`: one transient 400** (*An unexpected error occurred*) adding `pa_amount` (Decimal) right after the previous column. The identical payload worked when retried. Not reproduced, so not changed. A retry on 400/"unexpected error" for metadata POSTs may be worth adding.
3. **`pa app init` did not write a solution to `power.config.json`.** The app was pushed with `pa app push --solution-id <FlowOptimizationLab solutionid>` and landed in the solution.
4. **`pa` on headless Linux needs a Secret Service keyring.** `pa auth login` fails with `CachePersistenceError … GnomeKeyringError` because `usePlaintextFileOnLinux` is hard-coded to false. Workaround: run `dbus-daemon` + `gnome-keyring-daemon --unlock`, then `pa auth login --device-code`.
5. **The bound SharePoint connection was broken.** The existing connection (created 2025-07) had status *Error / Unauthorized*: `AADSTS50173 grant expired/revoked`. Binding it to the connection reference succeeded anyway, so a bound reference does not mean a healthy connection. The user fixed it in the portal. Worth a preflight check (`statuses[].status == Connected`) in `00-preflight` / `02-create-poc-flows`.
6. **Script 03, first run: `ECONNRESET`** at step 7/8, caused by a network proxy tunnel recycle in the container, not Dataverse. Rerun succeeded. The interrupted run left 799 rows, which were deleted afterwards with parallel single DELETEs (8 at a time, 0 failures, 0 remaining). The script only prints results at the very end and only cleans up on success, so writing partial results and cleaning up in a `finally` would make it more robust.

## App export

_Pending._

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

_Pending: runs after the app tests._
