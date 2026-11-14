# Power Automate Optimization Techniques — Catalog

Session: **From Hours to Seconds: Power Automate Optimization Techniques That Actually Work** (24 Nov 2026, 50 min)

Legend
- **In abstract**: ✅ promised in the abstract · ➕ extra technique we could add (bonus / backup demo / Q&A)
- **License**: 🟢 standard connectors only · 🟣 needs Premium (Dataverse, HTTP with Entra ID, Process plan…)
- **Impact**: rough order of magnitude we expect. **All timings must be measured on our own tenant before we put them on slides.**

---

## Platform facts worth putting on a slide

These numbers come from Microsoft Learn (checked Nov 2026). Most of the techniques below are justified by them.

| Limit | Value | Why it matters |
|---|---|---|
| Apply to each, concurrency | default **1** (sequential), configurable 1–50 | Out of the box, every loop is serial |
| Apply to each, max items | 5,000 (Low profile) / 100,000 (others) | Large datasets break loops |
| Pagination, max items | 5,000 (Low) / 100,000 (others) | Pagination is off by default; SharePoint Get items returns **100** rows unless you turn it on |
| Do until | default 60 iterations, max 5,000 | |
| Actions per flow definition | 500; nesting depth 8; variables 250 | |
| Expression length | 8,192 characters | Limits how big a single xpath/Select expression can get |
| Synchronous outbound/inbound request | **120 s** | Relevant to Power Apps / Code App calls and HTTP responses |
| Run duration | 30 days | |
| Power Platform requests / 24 h | **40,000** per paid user (Premium); **250,000** per Process / per-flow license | **Every action execution counts**, including each loop iteration |
| SharePoint connector | **600 calls / 60 s per connection** | Loops of Create item get throttled quickly |
| Dataverse service protection (per user, 5-min window) | 6,000 requests · 20 min execution time · 52 concurrent | Concurrency 50 against Dataverse runs into this |
| Dataverse Web API `$batch` | up to **1,000** requests per batch, changesets are atomic | |
| Dataverse bulk messages | CreateMultiple / UpdateMultiple / UpsertMultiple (standard + elastic), DeleteMultiple (**elastic only**) | Standard tables: 100–1,000 records per call, all-or-nothing |

> Key message for the audience: **an action is billed and throttled the same way whether it does real work or not.** A loop over 5,000 rows with 3 actions inside costs 15,000+ requests, roughly a third of a user's daily allowance.

---

## 1. Kill unnecessary loops ✅ 🟢

| # | Anti-pattern | Optimized pattern | Notes / snippet |
|---|---|---|---|
| 1.1 | `Apply to each` + `Append to array variable` to reshape objects | **Select** (data operation) | One action, no matter how many rows. In advanced (text) mode you can build nested objects. |
| 1.2 | `Apply to each` + `Condition` to keep matching rows | **Filter array** | Advanced mode: `@and(equals(item()?['Status'],'Open'), greater(float(item()?['Amount']),1000))` |
| 1.3 | Loop to extract one column | **Select in text mode**: map `item()?['Email']` → array of strings | Then `join(body('Select'),';')` gives a recipient list with no loop |
| 1.4 | Loop to add or rename one property | `addProperty(item(),'Total', mul(...))`, `setProperty`, `removeProperty` inside Select | Keeps every other column without listing them |
| 1.5 | Loop to remove duplicates | `union(variables('arr'), variables('arr'))` or `union(body('Select'), body('Select'))` | Distinct values in one expression |
| 1.6 | Loop to find common or missing items between two lists | `intersection()`; Filter array with `contains(body('SelectIdsB'), item()?['Id'])` / `not(contains(...))` | Diffing two systems with no loop |
| 1.7 | Loop to sort | `sort(array, 'Property')` + `reverse()` (built-in expressions) | Avoid Office Scripts or JS just to sort |
| 1.8 | Loop to take the first or last match | `first(body('Filter_array'))`, `last()`, `take()`, `skip()` | |
| 1.9 | Loop to build an HTML or CSV report | **Create HTML table / Create CSV table** from a Select | |
| 1.10 ➕ | **Nested loops for joins (O(n·m))**, e.g. look up a customer name for every order | **Dictionary / lookup-object pattern**: build one JSON object keyed by ID with Select + `join` + `json(concat('{', join(body('SelectKV'),','), '}'))`, then inside the main Select use `outputs('Dict')?[item()?['CustomerId']]` | Usually the biggest single gain in "enterprise" flows. Turns n·m iterations into 2 actions. |
| 1.11 ➕ | Loop to generate test data or date ranges | `range(0, 5000)` + Select | Also handy for seeding demo data |
| 1.12 ➕ | `Parse JSON` on huge payloads only for dynamic content | Reference directly: `body('HTTP')?['value']` | Parse JSON on a large schema adds validation time and makes the flow fragile |
| 1.13 ➕ | Loop to chunk records for batch calls | `chunk(array, 1000)` | Feeds the bulk patterns in section 3 |

**Decision framework (takeaway slide)**
1. Do you need a **side effect per item** (send an email, call an API that can't take a batch)? → you need a loop. Make it **concurrent** and keep it **thin**.
2. Are you only **reshaping, filtering, joining or aggregating data**? → no loop: Select, Filter array, expressions, xpath.
3. Does the target system **accept many records per call**? → no per-item loop: `$batch` or bulk messages (`chunk()` + a loop over the *chunks*).
4. Can the **source** do the work (filter, aggregate, expand)? → push it down to OData or FetchXML.

---

## 2. Aggregate without iterating: the XPath trick ✅ 🟢

Convert the array to XML once, then use XPath functions.

```text
// Compose "XmlRoot" (the JSON must have a single root)
xml(json(concat('{"root":{"r":', string(body('Select_Amounts')), '}}')))

Sum      : xpath(outputs('XmlRoot'), 'sum(/root/r/Amount)')
Count    : xpath(outputs('XmlRoot'), 'count(/root/r)')
Count if : xpath(outputs('XmlRoot'), 'count(/root/r[Status="Open"])')
Sum if   : xpath(outputs('XmlRoot'), 'sum(/root/r[Region="EMEA"]/Amount)')
Max      : xpath(outputs('XmlRoot'), 'number(/root/r/Amount[not(. < /root/r/Amount)][1])')   // XPath 1.0 idiom, O(n²): fine for hundreds, slow for 10k+
Min      : xpath(outputs('XmlRoot'), 'number(/root/r/Amount[not(. > /root/r/Amount)][1])')
Average  : div(xpath(...,'sum(...)'), xpath(...,'count(...)'))
```

Gotchas to call out live
- Property names must be valid XML element names: no spaces, can't start with a digit. Use **Select** first to rename and keep only the columns you need (this also makes the xml() conversion faster).
- XPath `sum()` returns a float, so round it with `formatNumber()` or `div(mul(x,100),100)`.
- An empty array produces an `<r/>` element, so guard for it with `empty()`.
- `max()` / `min()` exist as expressions and accept **an array of numbers**: `max(body('Select_Amounts_as_numbers'))` works if the Select (text mode) returns a flat number array. This is the better choice for Max/Min on large arrays. The XPath idiom is O(n²). Show both and measure.
- **Group-by without loops** ➕: `union()` gives the distinct keys → a Select over the keys with `xpath(outputs('XmlRoot'), concat('sum(/root/r[Region="', item(), '"]/Amount)'))` → one action gives the totals per group.

---

## 3. Go bulk or go home ✅

### 3a. SharePoint `$batch` 🟢 (Send an HTTP request to SharePoint is a standard connector)

- `POST {site}/_api/$batch`, `Content-Type: multipart/mixed; boundary=batch_<guid>`
- Inside it, a **changeset** (`multipart/mixed; boundary=changeset_<guid>`) holds the POST/PATCH/DELETE operations. GETs go outside changesets.
- Build the body with **Select** (one MIME part per record) + `join()`. **No loop needed to build the payload.**
- Chunk with `chunk(items, N)`. The limit isn't clearly documented: the PnP SDKs use **100 per batch**, and community reports go up to 1,000. **Test 100 / 500 / 1,000 on our tenant and show the curve.** Also watch the overall payload size.
- Parse the response with `split()` on the batch boundary, or count `HTTP/1.1 20` occurrences to get success/fail counts without a loop.
- Note for the audience: SharePoint changesets are **not truly transactional** the way Dataverse changesets are. Each operation can succeed or fail on its own, so present SP "changesets" carefully. **To verify on our tenant before the session.**
- ➕ Use `odata=nometadata` in the Accept header for smaller payloads everywhere.

### 3b. Dataverse `$batch` 🟣

- `POST [org]/api/data/v9.2/$batch`, up to **1,000** requests per batch. A **changeset** is atomic: one failure rolls back the whole changeset.
- `Prefer: odata.continue-on-error` → the batch continues and reports per-request errors (outside changesets).
- `Content-ID` + `$1` references let you create a parent and its children in one transactional call.
- Call it from a flow with **HTTP with Microsoft Entra ID (preauthorized)** 🟣, or the plain HTTP action + app registration 🟣.

### 3c. Dataverse bulk operation messages 🟣

| Message | Tables | Endpoint (Web API) | Response |
|---|---|---|---|
| CreateMultiple | standard + elastic | `POST /api/data/v9.2/<entityset>/Microsoft.Dynamics.CRM.CreateMultiple` body `{"Targets":[{...,"@odata.type":"Microsoft.Dynamics.CRM.<logicalname>"}]}` | 200 + `Ids` |
| UpdateMultiple | standard + elastic | `.../Microsoft.Dynamics.CRM.UpdateMultiple` (each target includes its primary key) | 204 |
| UpsertMultiple | standard + elastic | `.../Microsoft.Dynamics.CRM.UpsertMultiple` (`@odata.id` with alternate key) | 204 |
| DeleteMultiple | **elastic only** | `.../Microsoft.Dynamics.CRM.DeleteMultiple` | 204 |
| BulkDelete (async job) | standard | `BulkDelete` action with a QueryExpression | async system job |

- Recommended: **100–1,000 records per call** on standard tables. The whole call is **transactional** (all-or-nothing) on standard tables; elastic tables allow partial success.
- Each bulk call counts as **1** service-protection request, but **each target counts as one Power Platform request** for entitlement. Be honest about this on the slide: you save on throttling and time, not on PPR entitlement.
- From a flow: the Dataverse connector's **Perform a bound action** / **Perform an unbound action**. **Check whether CreateMultiple is exposed in the action picker for our table.** If it isn't, fall back to HTTP with Entra ID.
- ➕ For bulk deletes on standard tables, compare three options: a loop of Delete row, `$batch` DELETEs, and the `BulkDelete` async job.

### 3d. ➕ Other bulk surfaces

- **Microsoft Graph `$batch`**: 20 requests per batch (users, Planner, Teams, Outlook).
- **Excel + Office Scripts**: the Excel connector's "Add a row into a table" in a loop is one of the slowest patterns in the platform. Send the whole array to a **Run script** action that writes it with `range.setValues()` in a single call. This is often the most dramatic before/after on stage (minutes → seconds). 🟢 (Office Scripts is standard)
- **SQL**: a stored procedure with a JSON/TVP parameter instead of an Insert row loop. 🟣

---

## 4. Filter at the source ✅

| # | Anti-pattern | Optimized | Connector |
|---|---|---|---|
| 4.1 | Get items (all) → Filter array / Condition | **Filter Query** (OData) `Status eq 'Open' and Amount gt 1000` | SP 🟢 / DV 🟣 |
| 4.2 | Retrieve every column | **$select**: SP "Limit columns by view" or `Send HTTP` with `$select`; DV "Select columns" | Payload drops by 80–95% on wide lists |
| 4.3 | Get item inside a loop for each lookup | **$expand**: SP `$select=Title,Customer/Title&$expand=Customer`; DV "Expand Query" `primarycontactid($select=fullname)` | Removes a nested loop entirely |
| 4.4 | Count rows by retrieving everything | DV "Row count" / `$count=true`; SP `ItemCount` from list metadata | One call |
| 4.5 ➕ | Aggregate in the flow | **Dataverse FetchXML aggregates** (`aggregate="true"`, sum/avg/count/groupby) or OData `$apply=groupby((region),aggregate(amount with sum as total))` | The server does the math |
| 4.6 | Pagination on with a huge threshold "just in case" | Right-size the **pagination threshold**, use `$top` and order + filter | Pagination on SP Get items **defaults to off → 100 items** (a correctness bug as much as a performance issue) |
| 4.7 ➕ | Full re-read every run | **Delta / change-based**: SP "Get changes for an item or a file (properties only)", Graph delta, Dataverse change tracking, or a "Modified gt @{lastRun}" filter | |
| 4.8 | Flow runs for every change and exits on a Condition | **Trigger conditions** (`@equals(triggerOutputs()?['body/Status/Value'],'Approved')`), Dataverse "Filter rows" + "Select columns" on the trigger | Zero runs, zero billed actions when nothing is relevant |
| 4.9 ➕ | Get item right after the trigger | Use the trigger body (it already has the fields) | Saves one call per run |
| 4.10 ➕ | Indexed columns missing on SP lists > 5,000 | Index the columns used in Filter Query | Otherwise you hit the list view threshold |
| 4.11 ➕ | Dataverse List rows paginated via Do until | Use the built-in pagination, or `$skiptoken` handling plus a FetchXML paging cookie only when needed | |

---

## 5. Tune the platform itself ✅

| # | Topic | Guidance |
|---|---|---|
| 5.1 | **Concurrency on Apply to each** | Default 1. Set 20–50 for independent per-item work. **It bites you when**: you use `Set variable`/`Append` inside (race conditions, wrong results), the target throttles (SP 600/min, DV 52 concurrent), or order matters. **Demo the race condition**: concurrent loop + Set variable → wrong output. |
| 5.2 | **Compose vs variables** | Variables are global and locked during concurrent loops, and every Set/Increment is a billed action. Use Compose for values computed once and Select/expressions for collections. Use variables only for accumulators in sequential logic. |
| 5.3 | **Trigger concurrency** (flow-level) | Limit parallel runs (1–100) to protect downstream systems or guarantee order. Enabling it can't be undone on that flow. |
| 5.4 | **Split On** (debatching) on triggers | One run per item vs. one run for the batch: choose deliberately. |
| 5.5 ➕ | **Parallel branches** | Independent calls (get customers ‖ get products ‖ get prices) run in parallel, so total time = the slowest branch, not the sum. |
| 5.6 ➕ | **Child flows / fan-out** | Split 50k records into chunks and process them in parallel child flows (Run a Child Flow + concurrency) when a single run's loop limits or duration are the bottleneck. |
| 5.7 ➕ | **Retry policy** | Default exponential retries (up to 12 on Medium/High). Tune for idempotency; make sure batch calls handle 429 + `Retry-After`. |
| 5.8 ➕ | **Asynchronous response pattern** | Respond to the caller (Power Apps / HTTP) immediately with a run ID, then keep working. Avoids the 120 s timeout. We use this in the Code App harness. |
| 5.9 ➕ | **Performance profile** | Throughput depends on the license (Low/Medium/High). The Process plan gives High. Mention as context. |
| 5.10 | **Read run analytics** | Run history action durations (find the slow action, not the slow flow), Analytics tab, **Automation Center**, the Dataverse **flowrun** table (elastic, includes duration and status; 28-day default retention), and the Power Platform admin center request reports. |
| 5.11 ➕ | **Secure inputs/outputs + Static results** | Not a performance feature itself, but static results make before/after testing reproducible. |
| 5.12 ➕ | **Dataverse headers** 🟣 | `MSCRM.SuppressDuplicateDetection`, `MSCRM.BypassBusinessLogicExecution` (needs a privilege). Use with care and say so on the slide. |

---

## 6. Diagnose before you optimize ✅

1. Open the slowest run and sort/scan the action durations. Usually **one action or one loop** accounts for 90% of the time.
2. Expand loop iterations: check the **per-iteration duration × count**.
3. Check **throttling**: look for retries/429 in the action's raw outputs.
4. Check **run frequency**: too many runs → trigger conditions.
5. Measure **PPR consumption**: actions × runs. Compare against the 40k/day per user allowance.
6. Use the **flowrun** table (Dataverse) for history and trends, and Automation Center for failures and retries.

---

## Licensing caveat (affects the abstract wording)

The abstract says the techniques *"don't require premium licenses, custom code, or exotic workarounds"*. That is true for sections 1, 2, 3a, 4 (SharePoint), and 5. **Dataverse (3b, 3c, 4 DV) needs Premium**, because the Dataverse connector and HTTP with Entra ID are premium connectors. The **Code App benchmark harness also needs Premium** (code apps require Power Apps Premium / pay-as-you-go for end users). We don't need to change the abstract, but the deck should be explicit: *"SharePoint patterns = standard; Dataverse patterns = you already have premium if you use Dataverse."*
