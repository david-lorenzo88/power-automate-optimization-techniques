# Lab results

Measured on **PREMIUM-US** (Sandbox, US) on 14–15 Nov 2026, flows launched through `LAB-TestRunner`
(`lab/scripts/30-run.mjs`). Durations are the flows' own measurement (start of the flow → telemetry update),
so they include the same fixed overhead (telemetry row, respond early) for SLOW and FAST.
Regenerate the table with `node lab/scripts/40-results.mjs`.

## Slow vs fast pairs

| UC | Volume | SLOW | MID | FAST | Speed-up | Actions SLOW → FAST | Calls SLOW → FAST | Payload SLOW → FAST | Same result |
|---|---:|---:|---:|---:|---:|---|---|---|---|
| UC1 | 500 | 113.9 s |  | 0.9 s | **130×** | 664 → 17 | 1 → 1 | 52 KB → 52 KB | ✅ |
| UC2 | 200 | 3.0 min |  | 1.5 s | **124×** | 616 → 20 | 201 → 2 | 20 KB → 154 KB | ✅ |
| UC3 | 100 | 4.2 min |  | 1.4 s | **177×** | 947 → 20 | 1 → 1 | 10 KB → 10 KB | ✅ |
| UC4 | 100 | 75.0 s | 7.5 s | 18.7 s | **4.0×** | 117 → 24 | 101 → 2 | – → – | ✅ |
| UC4 | 1,000 | 11.1 min | 83.0 s | 44.9 s | **15×** | 1017 → 42 | 1001 → 5 | – → – | ✅ |
| UC5 | 100 | 28.8 s | 1.9 s | 3.4 s | **8.5×** | 117 → 19 | 101 → 2 | – → – | ✅ |
| UC5 | 1,000 | 5.6 min | 12.6 s | 10.4 s | **32×** | 1017 → 19 | 1001 → 2 | – → – | ✅ |
| UC6 | 2,000 | 59.6 s |  | 1.1 s | **53×** | 317 → 17 | 101 → 1 | 5.7 MB → 11 KB | ✅ |
| UC6DV | 5,000 | 8.8 s |  | 0.2 s | **36×** | 18 → 16 | 1 → 1 | 10.1 MB → 1 KB | ✅ |
| UC8 | 1,000 | 4.6 s |  | 3.7 s | **1.3×** | 216 → 17 | 1 → 1 | 288 KB → 2.8 MB | ❌ by design: 100\|12163.15 vs 1000\|499995.00 |
| UC9 | 20 | 49.2 s |  | 23.0 s | **2.1×** | 38 → 19 | 22 → 3 | – → – | ✅ |
| UC9 | 500 | 12.6 min |  | 17.8 s | **43×** | 518 → 19 | 502 → 3 | – → – | ✅ |
| UC10 | 200 | 76.0 s |  | 6.2 s | **12×** | 415 → 428 | 0 → 10 | – → – | ✅ |

Notes:
- **Actions** are the flows' own estimate (`pa_actionsestimated`); UC10 fan-out moves the actions into child flows rather than removing them (it saves wall-clock time, not requests).
- **UC2 payload** is higher for FAST because it reads all 2,000 customers once; SLOW makes 200 small calls instead.
- **UC4 at 100 items**: MID (concurrency 50) beat FAST. At 1,000 items the tuned FAST (250 per batch, 4 batches in parallel) wins. An earlier FAST version with one 500-item batch after another took 165.3 s and created only 949 of 1,000: SharePoint stops a batch after ~90 s of server time (see `poc/RESULTS.md` #8).
- **UC9 at 20 rows** is only 2×: each Office Script call costs a few seconds, so the gap only opens with volume.
- **UC8** is meant to fail: SLOW reads 100 of 1,000 rows (no pagination) and its total comes from a race on a variable in a concurrent loop.

## Other results

| What | Result |
|---|---|
| UC7 trigger conditions | 100 item changes (1 in 10 relevant) → **UC7-SLOW ran 100 times** (90 *Ignored*, 10 *Relevant*), **UC7-FAST ran 10 times** |
| UC5 rollback, 1 bad row | Loop (SLOW, MID): **19 of 20** rows created, status *Partial*. CreateMultiple (FAST): **0 of 100**, whole request rolled back |
| SEED-SP | Customers 2,000 + Orders 20,000 (26 columns) + UC7Items 100 via `$batch` (250–400 per batch, 4 in parallel): **16.6 min** |
| SEED-DV | `pa_order` 20,000 rows via CreateMultiple (1,000 per call, 4 in parallel): **49.8 s** |
| RESET uc5 | BulkDelete job submitted in 1.2 s (asynchronous; `RunNow: true` is refused for this table) |
| RESET uc7 | 100 UC7Items back to *Pending* in one `$batch` |
| RESET uc4 | 4,249 OrderImport items deleted with `$batch` DELETE (300 per batch, Do until empty): **6.0 min**, 0 left |
| RESET uc9 | Both UC9 workbooks cleared with the `Uc9Reset` Office Script: 8.8 s |

## Issues fixed while building

| Problem | Fix |
|---|---|
| `$batch` updates failed: *"The parameter __metadata does not exist in method GetById"* | SharePoint ignores `X-HTTP-Method: MERGE` inside a batch → request line `PATCH` |
| UC10 couldn't read child results | `result('loop')` returns one entry per action with an `outputs` array of iterations |
| UC10 / UC3 KPI step failed | `add()` takes exactly two arguments |
| SLOW loops with failures still showed *Succeeded* | Flows report *Partial* when fewer records were written than requested |
| BulkDelete 400: *"does not support Synchronous Bulk Delete"* | `RunNow: false` + typed `QueryExpression` |
| Deploy hung on one stalled request | Dataverse helper: 2-minute timeout + retry |
