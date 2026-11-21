# Limits and measurements behind the rules

## Platform limits (Microsoft Learn, Nov 2026)

| Limit | Value | Rule it drives |
|---|---|---|
| Apply to each concurrency | Default **1** (sequential), max 50 | BP006, BP005 |
| Apply to each items | 5,000 (Low profile) / 100,000 (others) | BP001, fan-out |
| Pagination | Off by default; max 5,000 (Low) / 100,000 items | BP009 |
| SharePoint *Get items* without pagination or `$top` | **100 items** (measured) | BP009 |
| Dataverse *List rows* default page | 5,000 rows | BP009 |
| Do until | Default 60 iterations / 1 hour; max 5,000 | BP016 |
| Actions per definition / nesting depth / variables | 500 / 8 / 250 | design |
| Expression length | 8,192 characters | xpath, Select |
| Synchronous request / Power Apps response | **120 s** | BP013 |
| Power Platform requests per 24 h | 40,000 per Premium user; 250,000 per Process licence | every rule |
| SharePoint connector | 600 calls / 60 s per connection | BP004, BP006 |
| Dataverse service protection (per user, 5-min window) | 6,000 requests, 20 min execution time, 52 concurrent | BP004, BP006 |
| Dataverse `$batch` | up to 1,000 requests; changesets atomic | pattern 6 |
| Dataverse bulk messages | CreateMultiple / UpdateMultiple / UpsertMultiple: 100–1,000 per call, standard tables all-or-nothing; DeleteMultiple elastic only | pattern 6 |
| Microsoft Graph `$batch` | 20 requests | — |

## Measured in the Flow Optimization Lab (PREMIUM-US sandbox, Nov 2026)

Same input, same checksum for every pair. Durations are the flows' own measurement.

| Use case | Volume | SLOW | FAST | Speed-up | Actions SLOW → FAST |
|---|---:|---:|---:|---:|---|
| Data shaping: loop + append → Filter array + Select | 500 | 113.9 s | 0.9 s | 130× | 664 → 17 |
| Join: query per row → dictionary | 200 | 3.0 min | 1.5 s | 124× | 616 → 20 |
| Aggregation: nested loops + variables → xpath | 100 | 4.2 min | 1.4 s | 177× | 947 → 20 |
| SharePoint create: item per row → `$batch` 250 × 4 parallel | 1,000 | 11.1 min | 44.9 s | 15× | 1,017 → 42 |
| Dataverse create: row per row → CreateMultiple | 1,000 | 5.6 min | 10.4 s | 32× | 1,017 → 19 |
| Filter at source: get all + filter + get per row → `$filter`/`$select`/`$expand` | 2,000 | 59.6 s | 1.1 s | 53× | 317 → 17 |
| Dataverse aggregate: list all + filter → FetchXML aggregate | 5,000 | 8.8 s | 0.2 s | 36× | payload 10.1 MB → 1 KB |
| Excel: Add a row per row → Office Script `setValues()` | 500 | 12.6 min | 17.8 s | 43× | 518 → 19 |
| Fan-out: sequential loop → chunk() + child flows | 200 | 76.0 s | 6.2 s | 12× | 415 → 428 (moved, not removed) |
| Trigger conditions | 100 changes, 10 relevant | 100 runs | 10 runs | — | — |

## Findings that are easy to get wrong

- **SharePoint `$batch` has a time limit, not an item limit.** Parts are processed serially (~95–170 ms each) and every part
  after ~90 s of server time returns `500 ClientServiceTimeoutException`. The outer response is still HTTP 200.
  1,000 items in one batch → 873–941 created. Keep ≤ 500 per batch.
- **SharePoint changesets are not atomic.** 100 creates in one changeset with the first one invalid → 99 created, 1 × 400.
- **Dataverse is atomic.** CreateMultiple with one invalid row → 0 of 100 created. A loop of *Add a new row* → 19 of 20 created.
- **Concurrency on a per-item loop helps, but batching wins at volume.** At 100 items, concurrency 50 (7.5 s) beat `$batch`
  (18.7 s). At 1,000 items `$batch` (44.9 s) beat concurrency 50 (83 s).
- **Responding late breaks the caller.** Power Apps `Run()` rejected at 120 s, and the run shows as Failed even though the work finished.
- **CreateMultiple is not reachable from the Dataverse connector.** It's bound to the entity set; *Perform a bound action* calls
  `{entitySet}({id})/{action}`. Use HTTP with Microsoft Entra ID.
- **`flowrun.duration` is in milliseconds** and new rows appear there after a few minutes.
