# Best-practice patterns (copy-ready definition JSON)

Every snippet comes from a flow that was deployed and measured (FAST variants in `lab/flows`), with tenant values
replaced by `<placeholders>`. Shapes follow the power-automate plugin: `OpenApiConnection`, slash-path parameter keys,
**no `authentication` block** on connector actions. Always confirm connector parameter names with `get_operation_details`.

`runAfter` is omitted below unless it matters. Add it as usual.

---

## 1. Filter array + Select instead of a loop (BP001)

Replaces *Apply to each → Condition → Compose → Append to array variable*. Measured: 500 orders, 113.9 s → 0.9 s (130×).

```json
"Open_orders": {
  "type": "Query",
  "inputs": {
    "from": "@body('Get_orders')?['value']",
    "where": "@equals(item()?['Status'], 'Open')"
  }
},
"Shaped": {
  "type": "Select",
  "inputs": {
    "from": "@body('Open_orders')",
    "select": {
      "orderId": "@item()?['Title']",
      "amount": "@item()?['Amount']",
      "amountWithVat": "@mul(item()?['Amount'], 1.21)"
    }
  }
}
```

- Keep every column and add one: `"select": "@addProperty(item(), 'Total', mul(item()?['Qty'], item()?['Price']))"`.
- One column as a flat array: `"select": "@item()?['Email']"`, then `join(body('Emails'), ';')`.
- Distinct values: `union(body('S'), body('S'))`. Sort: `sort(body('S'), 'Amount')`, `reverse(...)`.
- Items in A but not in B: Filter array over A with `@not(contains(body('Ids_B'), item()?['Id']))`.

## 2. Dictionary join instead of a lookup per row or a nested loop (BP002, BP003)

Read the lookup table **once**, turn it into a JSON object keyed by the join column, and look values up inside one Select.
Measured: 200 orders × 2,000 customers, 3.0 min / 201 calls → 1.5 s / 2 calls (124×).

```json
"Dict_parts": {
  "type": "Select",
  "inputs": {
    "from": "@body('Get_customers')?['value']",
    "select": "@concat('\"', item()?['Title'], '\":', string(item()))"
  }
},
"Dict": {
  "type": "Compose",
  "inputs": "@json(concat('{', join(body('Dict_parts'), ','), '}'))"
},
"Enriched": {
  "type": "Select",
  "inputs": {
    "from": "@body('Get_orders')?['value']",
    "select": {
      "orderId": "@item()?['Title']",
      "customerName": "@outputs('Dict')?[item()?['CustCode']]?['CustomerName']"
    }
  }
}
```

Keys must be unique strings. If the key can contain `"` or `\`, escape it, or key on an ID.
Alternative when the data is in one SharePoint list with a lookup column: `$expand` (pattern 4).

## 3. Aggregate with xml() + xpath(), group by with union() (BP001)

Replaces loops with *Increment variable* / *Set variable*. Measured: 100 orders × 5 regions, 4.2 min → 1.4 s (177×).

```json
"Orders_xml": {
  "type": "Compose",
  "inputs": "@xml(json(concat('{\"r\":{\"o\":', string(body('Get_orders')?['value']), '}}')))"
},
"Region_list": { "type": "Select", "inputs": { "from": "@body('Get_orders')?['value']", "select": "@item()?['Region']" } },
"Regions": { "type": "Compose", "inputs": "@sort(union(body('Region_list'), body('Region_list')))" },
"Summary": {
  "type": "Select",
  "inputs": {
    "from": "@outputs('Regions')",
    "select": {
      "region": "@item()",
      "count": "@int(xpath(outputs('Orders_xml'), concat('count(/r/o[Region=\"', item(), '\"])')))",
      "sum": "@float(xpath(outputs('Orders_xml'), concat('sum(/r/o[Region=\"', item(), '\"]/Amount)')))"
    }
  }
}
```

- Totals without grouping: `xpath(outputs('Orders_xml'), 'sum(/r/o/Amount)')`, `count(/r/o)`, `count(/r/o[Status="Open"])`.
- Max/min: `max(json(concat('[', join(xpath(outputs('Orders_xml'), '/r/o/Amount/text()'), ','), ']')))`, or `max()` on a Select that returns a number array.
- Element names must be valid XML. Rename columns with a Select first if they contain spaces or start with a digit.

## 4. Filter, trim and page at the source (BP007, BP008, BP009, BP010)

SharePoint *Get items* with a filter, a column view and explicit pagination:

```json
"Get_items": {
  "type": "OpenApiConnection",
  "inputs": {
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connectionName": "shared_sharepointonline", "operationId": "GetItems" },
    "parameters": {
      "dataset": "<site-url>",
      "table": "<list-guid>",
      "$filter": "Status eq 'Open' and Amount gt 1000",
      "view": "<view-guid-with-only-the-needed-columns>",
      "$top": 5000
    }
  },
  "runtimeConfiguration": { "paginationPolicy": { "minimumItemCount": 20000 } }
}
```

SharePoint REST through *Send an HTTP request to SharePoint* when you also need `$select` + `$expand` (no Get item per row).
Measured: 2,000 orders, 59.6 s / 5.7 MB → 1.1 s / 11 KB (53×).

```json
"Get_orders": {
  "type": "OpenApiConnection",
  "inputs": {
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connectionName": "shared_sharepointonline", "operationId": "HttpRequest" },
    "parameters": {
      "dataset": "<site-url>",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists(guid'<list-guid>')/items?$select=Title,Amount,Customer/CustomerName&$expand=Customer&$filter=Region eq 'EMEA' and Status eq 'Open'&$top=5000",
      "parameters/headers": { "Accept": "application/json;odata=nometadata" }
    }
  }
}
```

`odata=nometadata` keeps the payload small. Follow `odata.nextLink` if more than 5,000 rows can match.

Dataverse *List rows*: `$select`, `$filter`, `$top`, or let the server aggregate with FetchXML.
Measured: 5,000 rows, 8.8 s / 10.1 MB → 0.2 s / 1 KB (36×).

```json
"Aggregate": {
  "type": "OpenApiConnection",
  "inputs": {
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps", "connectionName": "shared_commondataserviceforapps", "operationId": "ListRecords" },
    "parameters": {
      "entityName": "<entity-set-name>",
      "fetchXml": "<fetch aggregate=\"true\"><entity name=\"<logical-name>\"><attribute name=\"<id-column>\" alias=\"n\" aggregate=\"count\"/><attribute name=\"<amount-column>\" alias=\"s\" aggregate=\"sum\"/><filter><condition attribute=\"<status-column>\" operator=\"eq\" value=\"Open\"/></filter></entity></fetch>"
    }
  }
}
```

## 5. SharePoint $batch: many items per request (BP004, BP015)

Loop over **chunks**, build each batch body with Select + join (no loop per item), send it with
*Send an HTTP request to SharePoint*. Keep **≤ 500 items per batch** (SharePoint stops after ~90 s of server time).
250 per batch × 4 batches in parallel was fastest. Measured: 1,000 items, 11.1 min → 44.9 s (15×).

```json
"CRLF": { "type": "Compose", "inputs": "@decodeUriComponent('%0D%0A')" },
"Chunks": { "type": "Compose", "inputs": "@chunk(body('Rows_to_create'), 250)" },
"Batch_loop": {
  "type": "Foreach",
  "foreach": "@outputs('Chunks')",
  "runtimeConfiguration": { "concurrency": { "repetitions": 4 } },
  "actions": {
    "B_ids": { "type": "Compose", "inputs": { "batch": "@{guid()}", "cs": "@{guid()}" } },
    "B_parts": {
      "type": "Select",
      "inputs": {
        "from": "@items('Batch_loop')",
        "select": "@concat('--changeset_', outputs('B_ids')?['cs'], outputs('CRLF'), 'Content-Type: application/http', outputs('CRLF'), 'Content-Transfer-Encoding: binary', outputs('CRLF'), outputs('CRLF'), 'POST <site-url>/_api/web/lists(guid''<list-guid>'')/items HTTP/1.1', outputs('CRLF'), 'Content-Type: application/json;odata=verbose', outputs('CRLF'), outputs('CRLF'), concat('{\"__metadata\":{\"type\":\"<ListItemEntityTypeFullName>\"},\"Title\":\"', item()?['Title'], '\",\"Amount\":', string(item()?['Amount']), '}'), outputs('CRLF'), outputs('CRLF'))"
      },
      "runAfter": { "B_ids": ["Succeeded"] }
    },
    "B_body": {
      "type": "Compose",
      "inputs": "@concat('--batch_', outputs('B_ids')?['batch'], outputs('CRLF'), 'Content-Type: multipart/mixed; boundary=\"changeset_', outputs('B_ids')?['cs'], '\"', outputs('CRLF'), 'Content-Transfer-Encoding: binary', outputs('CRLF'), outputs('CRLF'), join(body('B_parts'), ''), '--changeset_', outputs('B_ids')?['cs'], '--', outputs('CRLF'), outputs('CRLF'), '--batch_', outputs('B_ids')?['batch'], '--', outputs('CRLF'))",
      "runAfter": { "B_parts": ["Succeeded"] }
    },
    "B_send": {
      "type": "OpenApiConnection",
      "inputs": {
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connectionName": "shared_sharepointonline", "operationId": "HttpRequest" },
        "parameters": {
          "dataset": "<site-url>",
          "parameters/method": "POST",
          "parameters/uri": "_api/$batch",
          "parameters/headers": { "Content-Type": "multipart/mixed; boundary=\"batch_@{outputs('B_ids')?['batch']}\"", "Accept": "application/json;odata=verbose" },
          "parameters/body": "@outputs('B_body')"
        }
      },
      "runAfter": { "B_body": ["Succeeded"] }
    },
    "B_text": {
      "type": "Compose",
      "inputs": "@if(startsWith(string(body('B_send')), '{\"$content'), base64ToString(json(string(body('B_send')))?['$content']), string(body('B_send')))",
      "runAfter": { "B_send": ["Succeeded"] }
    }
  }
}
```

- The outer response is HTTP 200 even when parts fail. Count parts: `sub(length(split(outputs('B_text'), 'HTTP/1.1 201')), 1)`,
  and the same for `'HTTP/1.1 4'` and `'HTTP/1.1 5'`. Report the flow as Partial when created < requested.
- Get `<ListItemEntityTypeFullName>` once with `GET _api/web/lists(guid'<list-guid>')?$select=ListItemEntityTypeFullName`.
- Updates: request line `PATCH …/items(<id>)` plus header `IF-MATCH: *`. Deletes: `DELETE …/items(<id>)`.
- SharePoint changesets are **not** atomic: with one bad item, the other 99 were still created.

## 6. Dataverse CreateMultiple / UpdateMultiple / UpsertMultiple (BP004)

The bulk messages are bound to the entity **set**, so the Dataverse connector's *Perform a bound action* can't call them.
Use **HTTP with Microsoft Entra ID (preauthorized)** (`shared_webcontents`, base resource URL and resource URI = the org URL).
100–1,000 records per call. On standard tables the whole call is all-or-nothing.
Measured: 1,000 rows, 5.6 min → 10.4 s (32×); with one bad row: 0 of 100 written (loop: 19 of 20).

```json
"Chunks": { "type": "Compose", "inputs": "@chunk(body('Rows_to_create'), 1000)" },
"Bulk_loop": {
  "type": "Foreach",
  "foreach": "@outputs('Chunks')",
  "actions": {
    "Targets": {
      "type": "Select",
      "inputs": {
        "from": "@items('Bulk_loop')",
        "select": "@concat('{\"@odata.type\":\"Microsoft.Dynamics.CRM.<logical-name>\",\"<name-column>\":\"', item()?['Name'], '\",\"<amount-column>\":', string(item()?['Amount']), '}')"
      }
    },
    "Create_multiple": {
      "type": "OpenApiConnection",
      "inputs": {
        "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_webcontents", "connectionName": "shared_webcontents", "operationId": "InvokeHttp" },
        "parameters": {
          "request/method": "POST",
          "request/url": "<org-url>/api/data/v9.2/<entity-set-name>/Microsoft.Dynamics.CRM.CreateMultiple",
          "request/headers": { "Content-Type": "application/json; charset=utf-8", "Accept": "application/json", "OData-Version": "4.0", "OData-MaxVersion": "4.0" },
          "request/body": "@{concat('{\"Targets\":[', join(body('Targets'), ','), ']}')}"
        }
      },
      "runAfter": { "Targets": ["Succeeded"] }
    }
  }
}
```

- `UpdateMultiple`: each target includes its primary key. `UpsertMultiple`: each target has `"@odata.id": "<entity-set-name>(<alt-key>='<value>')"`.
- `DeleteMultiple` is elastic tables only. For standard tables use the asynchronous `BulkDelete` job (`RunNow: false`) or `$batch` DELETE.
- A mix of operations, or parent + children in one transaction: Dataverse connector *Perform a changeset request* (`ExecuteChangeset`).

## 7. Excel: one Office Script call instead of Add a row per row (BP004)

Measured: 500 rows, 12.6 min → 17.8 s (43×). Build the rows with Select, pass them as one JSON string.

```json
"Rows": {
  "type": "Select",
  "inputs": { "from": "@body('Source')", "select": "@createArray(item()?['OrderId'], item()?['Region'], item()?['Amount'])" }
},
"Write_rows": {
  "type": "OpenApiConnection",
  "inputs": {
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_excelonlinebusiness", "connectionName": "shared_excelonlinebusiness", "operationId": "RunScriptProdV2" },
    "parameters": {
      "source": "<site-url or 'me'>", "drive": "<drive-id>", "file": "<file-id>",
      "scriptSource": "<site-url or 'me'>", "scriptDrive": "<drive-id>", "scriptId": "<script-file-id>",
      "ScriptParameters/rowsJson": "@{string(body('Rows'))}"
    }
  }
}
```

```ts
function main(workbook: ExcelScript.Workbook, rowsJson: string): number {
  const rows: (string | number)[][] = JSON.parse(rowsJson);
  const table = workbook.getTable("Orders");
  if (rows.length === 0) return table.getRowCount();
  const sheet = table.getWorksheet();
  const header = table.getHeaderRowRange();
  const existing = table.getRowCount();
  sheet.getRangeByIndexes(header.getRowIndex() + 1 + existing, header.getColumnIndex(), rows.length, rows[0].length).setValues(rows);
  table.resize(sheet.getRangeByIndexes(header.getRowIndex(), header.getColumnIndex(), 1 + existing + rows.length, rows[0].length));
  return existing + rows.length;
}
```

Each script call costs a few seconds, so the gap only opens with volume (20 rows: 2×; 500 rows: 43×).

## 8. Trigger conditions instead of Condition + Terminate (BP011)

Measured: 100 item changes, 10 relevant → 100 runs became 10 runs.

```json
"When_an_item_is_created_or_modified": {
  "type": "OpenApiConnection",
  "recurrence": { "interval": 1, "frequency": "Minute" },
  "splitOn": "@triggerOutputs()?['body/value']",
  "inputs": {
    "host": { "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline", "connectionName": "shared_sharepointonline", "operationId": "GetOnUpdatedItems" },
    "parameters": { "dataset": "<site-url>", "table": "<list-guid>" }
  },
  "conditions": [
    { "expression": "@equals(triggerOutputs()?['body/Status/Value'], 'Approved')" }
  ]
}
```

- All conditions must be true (AND). For OR, put one `@or(...)` expression in a single condition.
- A choice column is `body/<Column>/Value`, a text column `body/<Column>`.
- Dataverse triggers: use *Filter rows* (`subscriptionRequest/filterexpression`) and *Select columns* (`subscriptionRequest/filteringattributes`) on the trigger itself.

## 9. Respond early to callers that wait (BP013)

Power Apps / HTTP / Copilot callers fail at 120 s. Respond first, then do the work, and report progress through a table.

```json
"Respond_to_app": {
  "type": "Response",
  "kind": "PowerApp",
  "inputs": {
    "statusCode": 200,
    "body": { "runid": "@{workflow()?['run']?['name']}" },
    "schema": { "type": "object", "properties": { "runid": { "title": "runId", "type": "string", "x-ms-dynamically-added": true } } }
  },
  "runAfter": { "Create_status_row": ["Succeeded"] }
},
"Work": { "type": "Scope", "actions": { }, "runAfter": { "Respond_to_app": ["Succeeded"] } }
```

Run link for the status row: `concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])`.
(The `/solutions/~preferred/` form and `/edit` didn't open from a Code App in our tests; use `/flows/{id}/details` to reach the editor.)

## 10. Thin, concurrent loops for real side effects (BP005, BP006)

```json
"Notify_each": {
  "type": "Foreach",
  "foreach": "@body('Recipients')",
  "runtimeConfiguration": { "concurrency": { "repetitions": 20 } },
  "actions": { "Send_email": { "type": "OpenApiConnection", "inputs": { } } }
}
```

No `Set variable` inside. Compute per-item values with `item()` or a Compose. Collect results after the loop.

## 11. Fan-out to child flows for very large jobs

`chunk()` + Apply to each (concurrency 10) + *Run a Child Flow* per chunk. Measured: 200 records, 76.0 s → 6.2 s (12×).
This saves wall-clock time, not actions: the work moves into the child flows.

```json
"Chunks": { "type": "Compose", "inputs": "@chunk(body('Items'), 20)" },
"Fan_out": {
  "type": "Foreach",
  "foreach": "@outputs('Chunks')",
  "runtimeConfiguration": { "concurrency": { "repetitions": 10 } },
  "actions": {
    "Run_child": {
      "type": "Workflow",
      "inputs": { "host": { "workflowReferenceName": "<child-flow-id>" }, "body": { "text": "@{string(items('Fan_out'))}" } }
    }
  }
}
```

The child flow uses a manual (`kind: "Button"`) or Power Apps trigger and must end with a Response. Read the child results after
the loop with a Filter array over `result('Fan_out')` where `item()?['name']` equals `'Run_child'`, then take its `outputs`:
`result()` returns one entry per action, and each entry's `outputs` is an array with one element per iteration.
