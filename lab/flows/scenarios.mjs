// Slow-vs-fast flow pairs for docs/02-demo-scenarios.md (UC1–UC10) + seed/reset + test runner.
//
// buildLabFlows(cfg, ids) → [{ name, description, apis, meta, trigger, actions }]
//   cfg: lab/lab.config.json
//   ids: { '<flow name>': '<workflowid>' } — needed by flows that call child flows
//        (UC10-FAST, LAB-TestRunner). Without ids those flows are left out (deploy phase 1).
//   runnable: flows that are on; LAB-TestRunner only calls those.
//
// Every pair produces the same checksum from the same input (except UC8-SLOW, whose
// wrong checksum IS the demo). Data is deterministic (lib.mjs `data`), so the seed flows
// and the in-memory generators always agree.
import {
  ANY, CORR, CRLF, DV, HTTPDV, SP, VOL, XL, after, appendVar, benchmarkFlow, compose, cond, data, dvAction, dvHttp,
  filter, forEach, incVar, kpi, money, pad, pl, powerAppsTrigger, q, scope, select, setVar, spAction, spHttp, xlAction, xmlOf, xsum,
} from './lib.mjs';

const eq = (a, b) => ({ equals: [a, b] });
const and = (...xs) => ({ and: xs });
const BASE = 14; // header + footer actions every benchmark flow has (estimate for pa_actionsestimated)

export function buildLabFlows(cfg, ids = {}, { runnable } = {}) {
  const P = cfg.prefix;
  const site = cfg.siteUrl;
  const L = cfg.lists;
  const listUri = (name) => `_api/web/lists(guid'${L[name].id}')`;
  const flows = [];
  const add = (f) => flows.push(f);

  // ---------------------------------------------------------------- shared pieces
  // Orders 1..N with only the columns we need (identical step in SLOW and FAST of UC1–UC3).
  const getOrders = (runAfter = {}) => spHttp(cfg, 'GET',
    `${listUri('Orders')}/items?$select=Id,Title,CustCode,Region,Status,Amount&$filter=Id le @{${VOL}}&$orderby=Id&$top=5000`, runAfter);
  const ORDERS = "body('Get_orders')?['value']";
  const N_ORDERS = `length(${ORDERS})`;
  const ordersPayload = "length(string(body('Get_orders')))";

  // SharePoint $batch (one changeset) built with Select + join: actions <p>_ids/_parts/_body/_send/_text
  const spBatch = (p, itemsExpr, requestExpr, runAfter) => ({
    [`${p}_ids`]: compose({ batch: '@{guid()}', cs: '@{guid()}' }, runAfter),
    [`${p}_parts`]: select(itemsExpr,
      `@concat('--changeset_', outputs('${p}_ids')?['cs'], ${CRLF}, 'Content-Type: application/http', ${CRLF}, 'Content-Transfer-Encoding: binary', ${CRLF}, ${CRLF}, ${requestExpr}, ${CRLF})`,
      after(`${p}_ids`)),
    [`${p}_body`]: compose(
      `@concat('--batch_', outputs('${p}_ids')?['batch'], ${CRLF}, 'Content-Type: multipart/mixed; boundary="changeset_', outputs('${p}_ids')?['cs'], '"', ${CRLF}, 'Content-Transfer-Encoding: binary', ${CRLF}, ${CRLF}, join(body('${p}_parts'), ''), '--changeset_', outputs('${p}_ids')?['cs'], '--', ${CRLF}, ${CRLF}, '--batch_', outputs('${p}_ids')?['batch'], '--', ${CRLF})`,
      after(`${p}_parts`)),
    [`${p}_send`]: spHttp(cfg, 'POST', '_api/$batch', after(`${p}_body`), {
      headers: { 'Content-Type': `multipart/mixed; boundary="batch_@{outputs('${p}_ids')?['batch']}"`, Accept: 'application/json;odata=verbose' },
      body: `@outputs('${p}_body')`,
    }),
    // The connector returns multipart responses as {"$content": base64}; turn it into text.
    [`${p}_text`]: compose(
      `@if(startsWith(string(body('${p}_send')), '{"$content'), base64ToString(json(string(body('${p}_send')))?['$content']), string(body('${p}_send')))`,
      after(`${p}_send`, ['Succeeded', 'Failed'])),
  });
  const countIn = (textExpr, needle) => `sub(length(split(${textExpr}, ${q(needle)})), 1)`;
  const createReq = (listName, jsonExpr) =>
    `${q(`POST ${site}/${listUri(listName)}/items HTTP/1.1`)}, ${CRLF}, 'Content-Type: application/json;odata=verbose', ${CRLF}, ${CRLF}, ${jsonExpr}, ${CRLF}`;
  // Inside $batch SharePoint ignores X-HTTP-Method: the request line must say PATCH.
  const mergeReq = (listName, idExpr, jsonExpr) =>
    `${q(`PATCH ${site}/${listUri(listName)}/items(`)}, string(${idExpr}), ') HTTP/1.1', ${CRLF}, 'Content-Type: application/json;odata=verbose', ${CRLF}, 'IF-MATCH: *', ${CRLF}, ${CRLF}, ${jsonExpr}, ${CRLF}`;
  const deleteReq = (listName, idExpr) =>
    `${q(`DELETE ${site}/${listUri(listName)}/items(`)}, string(${idExpr}), ') HTTP/1.1', ${CRLF}, 'IF-MATCH: *', ${CRLF}`;
  const crlf = { CRLF: compose("@decodeUriComponent('%0D%0A')", {}) };

  // Dataverse aggregate (count + sum of amount) for the rows tagged with this run.
  const verifyImport = (runAfter) => dvAction('ListRecords', {
    entityName: `${P}_orderimports`,
    fetchXml: `<fetch aggregate="true"><entity name="${P}_orderimport"><attribute name="${P}_orderimportid" alias="n" aggregate="count"/><attribute name="${P}_amount" alias="s" aggregate="sum"/><filter><condition attribute="${P}_batchtag" operator="eq" value="@{${CORR}}"/></filter></entity></fetch>`,
  }, runAfter);
  const verifyImportChecksum = "concat(string(coalesce(first(body('Verify')?['value'])?['n'], 0)), '|', " + money("coalesce(first(body('Verify')?['value'])?['s'], 0)") + ')';
  const verifyImportCount = "int(coalesce(first(body('Verify')?['value'])?['n'], 0))";

  // ================================================================= UC1 Data shaping
  const shaped = (it) => ({ orderId: `@${it}?['Title']`, region: `@${it}?['Region']`, amount: `@${it}?['Amount']`, amountWithVat: `@mul(${it}?['Amount'], 1.21)` });
  const uc1Check = (arrExpr) => `concat(string(length(${arrExpr})), '|', ${money(xsum(arrExpr, 'amountWithVat'))})`;
  add(benchmarkFlow(P, {
    name: 'UC1-SLOW', scenario: 'UC1', variant: 'SLOW', technique: 'Apply to each + Condition + Compose + Append to array variable',
    description: 'UC1 SLOW: loop over orders, keep Open ones, reshape each and append it to an array variable.',
    apis: [SP], volumeDefault: 1000, volumeMax: 5000,
    vars: [{ name: 'Result', type: 'array', value: [] }],
    work: {
      Get_orders: getOrders(),
      Shape_loop: forEach(`@${ORDERS}`, {
        Is_open: cond(and(eq("@items('Shape_loop')?['Status']", 'Open')), {
          Shaped: compose(shaped("items('Shape_loop')"), {}),
          Append_shaped: appendVar('Result', "@outputs('Shaped')", after('Shaped')),
        }, {}, {}),
      }, after('Get_orders')),
      Kpi: kpi({
        processed: "length(variables('Result'))", retrieved: N_ORDERS, loops: N_ORDERS, calls: '1', payload: ordersPayload,
        actions: `add(${BASE + 2}, add(${N_ORDERS}, mul(2, length(variables('Result')))))`,
        checksum: uc1Check("variables('Result')"),
      }, after('Shape_loop', ['Succeeded', 'Failed'])),
    },
  }));
  add(benchmarkFlow(P, {
    name: 'UC1-FAST', scenario: 'UC1', variant: 'FAST', technique: 'Filter array + Select',
    description: 'UC1 FAST: Filter array (Status = Open) + Select reshapes all orders in two actions.',
    apis: [SP], volumeDefault: 1000, volumeMax: 5000,
    work: {
      Get_orders: getOrders(),
      Open_orders: filter(`@${ORDERS}`, "@equals(item()?['Status'], 'Open')", after('Get_orders')),
      Shaped: select("@body('Open_orders')", shaped('item()'), after('Open_orders')),
      Kpi: kpi({
        processed: "length(body('Shaped'))", retrieved: N_ORDERS, calls: '1', payload: ordersPayload, actions: String(BASE + 3),
        checksum: uc1Check("body('Shaped')"),
      }, after('Shaped')),
    },
  }));

  // ================================================================= UC2 Join / enrichment
  const enriched = (it, nameExpr, segExpr) => ({ orderId: `@${it}?['Title']`, custCode: `@${it}?['CustCode']`, amount: `@${it}?['Amount']`, customerName: `@${nameExpr}`, segment: `@${segExpr}` });
  const uc2Check = (arrExpr) => `concat(string(length(${arrExpr})), '|', ${money(xsum(arrExpr, 'amount'))}, '|', string(int(xpath(${xmlOf(arrExpr)}, 'count(/r/o[segment="Enterprise"])'))))`;
  add(benchmarkFlow(P, {
    name: 'UC2-SLOW', scenario: 'UC2', variant: 'SLOW', technique: 'Apply to each + Get customer per order',
    description: 'UC2 SLOW: for each order, query the Customers list for its customer (N connector calls).',
    apis: [SP], volumeDefault: 500, volumeMax: 5000,
    vars: [{ name: 'Result', type: 'array', value: [] }],
    work: {
      Get_orders: getOrders(),
      Join_loop: forEach(`@${ORDERS}`, {
        Get_customer: spHttp(cfg, 'GET', `${listUri('Customers')}/items?$select=Title,CustomerName,Segment&$filter=Title eq '@{items('Join_loop')?['CustCode']}'&$top=1`, {}),
        Enriched: compose(enriched("items('Join_loop')", "first(body('Get_customer')?['value'])?['CustomerName']", "first(body('Get_customer')?['value'])?['Segment']"), after('Get_customer')),
        Append_enriched: appendVar('Result', "@outputs('Enriched')", after('Enriched')),
      }, after('Get_orders')),
      Kpi: kpi({
        processed: "length(variables('Result'))", retrieved: N_ORDERS, loops: N_ORDERS, calls: `add(1, ${N_ORDERS})`, payload: ordersPayload,
        actions: `add(${BASE + 2}, mul(3, ${N_ORDERS}))`, checksum: uc2Check("variables('Result')"),
      }, after('Join_loop', ['Succeeded', 'Failed'])),
    },
  }));
  add(benchmarkFlow(P, {
    name: 'UC2-FAST', scenario: 'UC2', variant: 'FAST', technique: 'Dictionary pattern (Select + json(concat) + key lookup)',
    description: 'UC2 FAST: read all customers once, build a JSON dictionary keyed by customer code, enrich all orders in one Select.',
    apis: [SP], volumeDefault: 500, volumeMax: 5000,
    work: {
      Get_orders: getOrders(),
      Get_customers: spHttp(cfg, 'GET', `${listUri('Customers')}/items?$select=Title,CustomerName,Segment&$top=5000`, {}),
      Dict_parts: select("@body('Get_customers')?['value']", "@concat('\"', item()?['Title'], '\":', string(item()))", after('Get_customers')),
      Dict: compose("@json(concat('{', join(body('Dict_parts'), ','), '}'))", after('Dict_parts')),
      Enriched: select(`@${ORDERS}`, enriched('item()', "outputs('Dict')?[item()?['CustCode']]?['CustomerName']", "outputs('Dict')?[item()?['CustCode']]?['Segment']"), { Dict: ['Succeeded'], Get_orders: ['Succeeded'] }),
      Kpi: kpi({
        processed: "length(body('Enriched'))", retrieved: `add(${N_ORDERS}, length(body('Get_customers')?['value']))`, calls: '2',
        payload: `add(${ordersPayload}, length(string(body('Get_customers'))))`, actions: String(BASE + 6), checksum: uc2Check("body('Enriched')"),
      }, after('Enriched')),
    },
  }));

  // ================================================================= UC3 Aggregation
  const regionLine = (r, c, s, mn, mx) => `concat(${r}, ':', string(${c}), ':', ${money(s)}, ':', ${money(mn)}, ':', ${money(mx)})`;
  add(benchmarkFlow(P, {
    name: 'UC3-SLOW', scenario: 'UC3', variant: 'SLOW', technique: 'Nested loops + Increment / Set variable per region',
    description: 'UC3 SLOW: group by region with a loop over regions and an inner loop over all orders (count/sum/min/max in variables).',
    apis: [SP], volumeDefault: 300, volumeMax: 5000,
    vars: [
      { name: 'Results', type: 'array', value: [] }, { name: 'Cnt', type: 'integer', value: 0 }, { name: 'Sum', type: 'float', value: 0 },
      { name: 'MaxAmt', type: 'float', value: -1 }, { name: 'MinAmt', type: 'float', value: 1000000 },
    ],
    work: {
      Get_orders: getOrders(),
      Region_loop: forEach(`@createArray('EMEA', 'AMER', 'APAC', 'LATAM', 'MEA')`, {
        Reset_cnt: setVar('Cnt', 0, {}),
        Reset_sum: setVar('Sum', 0, after('Reset_cnt')),
        Reset_max: setVar('MaxAmt', -1, after('Reset_sum')),
        Reset_min: setVar('MinAmt', 1000000, after('Reset_max')),
        Order_loop: forEach(`@${ORDERS}`, {
          In_region: cond(and(eq("@items('Order_loop')?['Region']", "@items('Region_loop')")), {
            Add_count: incVar('Cnt', 1, {}),
            Add_sum: incVar('Sum', "@items('Order_loop')?['Amount']", after('Add_count')),
            Check_max: cond({ and: [{ greater: ["@float(items('Order_loop')?['Amount'])", "@variables('MaxAmt')"] }] }, {
              Set_max: setVar('MaxAmt', "@float(items('Order_loop')?['Amount'])", {}),
            }, {}, after('Add_sum')),
            Check_min: cond({ and: [{ less: ["@float(items('Order_loop')?['Amount'])", "@variables('MinAmt')"] }] }, {
              Set_min: setVar('MinAmt', "@float(items('Order_loop')?['Amount'])", {}),
            }, {}, after('Check_max')),
          }, {}, {}),
        }, after('Reset_min')),
        Has_orders: cond({ and: [{ greater: ["@variables('Cnt')", 0] }] }, {
          Append_region: appendVar('Results', { region: "@items('Region_loop')", count: "@variables('Cnt')", sum: "@variables('Sum')", min: "@variables('MinAmt')", max: "@variables('MaxAmt')" }, {}),
        }, {}, after('Order_loop')),
      }, after('Get_orders')),
      Summary: select("@sort(variables('Results'), 'region')", `@${regionLine("item()?['region']", "item()?['count']", "item()?['sum']", "item()?['min']", "item()?['max']")}`, after('Region_loop')),
      Kpi: kpi({
        processed: N_ORDERS, retrieved: N_ORDERS, loops: `add(5, mul(5, ${N_ORDERS}))`, calls: '1', payload: ordersPayload,
        actions: `add(${BASE + 3 + 5 * 6}, mul(9, ${N_ORDERS}))`, checksum: "join(body('Summary'), ';')",
      }, after('Summary')),
    },
  }));
  const xp = (path) => `xpath(outputs('Orders_xml'), concat('${path.replace('{R}', "', item(), '")}'))`;
  const nums = (path) => `json(concat('[', join(${xp(path)}, ','), ']'))`;
  add(benchmarkFlow(P, {
    name: 'UC3-FAST', scenario: 'UC3', variant: 'FAST', technique: 'xml() + xpath sum/count + min()/max(), group-by via union()',
    description: 'UC3 FAST: distinct regions with union(), then one Select computes count/sum (xpath) and min/max per region.',
    apis: [SP], volumeDefault: 300, volumeMax: 5000,
    work: {
      Get_orders: getOrders(),
      Region_list: select(`@${ORDERS}`, "@item()?['Region']", after('Get_orders')),
      Regions: compose("@sort(union(body('Region_list'), body('Region_list')))", after('Region_list')),
      Orders_xml: compose(`@${xmlOf(ORDERS)}`, after('Get_orders')),
      Summary: select("@outputs('Regions')", `@${regionLine('item()',
        `int(${xp('count(/r/o[Region="{R}"])')})`,
        xp('sum(/r/o[Region="{R}"]/Amount)'),
        `min(${nums('/r/o[Region="{R}"]/Amount/text()')})`,
        `max(${nums('/r/o[Region="{R}"]/Amount/text()')})`)}`, { Regions: ['Succeeded'], Orders_xml: ['Succeeded'] }),
      Kpi: kpi({
        processed: N_ORDERS, retrieved: N_ORDERS, calls: '1', payload: ordersPayload, actions: String(BASE + 6), checksum: "join(body('Summary'), ';')",
      }, after('Summary')),
    },
  }));

  // ================================================================= UC4 SharePoint bulk create
  const oiTitle = (i) => `concat(${CORR}, '-', string(${i}))`;
  const uc4Verify = (runAfter) => spHttp(cfg, 'GET', `${listUri('OrderImport')}/items?$select=Amount&$filter=BatchTag eq '@{${CORR}}'&$top=5000`, runAfter);
  const uc4Check = `concat(string(length(body('Verify')?['value'])), '|', ${money(xsum("body('Verify')?['value']", 'Amount'))})`;
  for (const [variant, conc] of [['SLOW', 1], ['MID', 50]]) {
    add(benchmarkFlow(P, {
      name: `UC4-${variant}`, scenario: 'UC4', variant,
      technique: conc === 1 ? 'Apply to each + Create item (concurrency 1)' : 'Apply to each + Create item (concurrency 50)',
      description: `UC4 ${variant}: one "Create item" call per row${conc > 1 ? ', 50 in parallel (watch for 429 throttling retries)' : ''}.`,
      apis: [SP], volumeDefault: 200, volumeMax: 5000,
      work: {
        Create_loop: forEach(`@range(0, ${VOL})`, {
          Create_item: spAction(cfg, 'PostItem', {
            table: L.OrderImport.id,
            'item/Title': `@{${oiTitle("items('Create_loop')")}}`,
            'item/Amount': `@${data.amount("add(items('Create_loop'), 1)")}`,
            'item/Region': `@{${data.region("items('Create_loop')")}}`,
            'item/BatchTag': `@{${CORR}}`,
          }, {}),
        }, {}, conc),
        Verify: uc4Verify(after('Create_loop', ['Succeeded', 'Failed'])),
        Kpi: kpi({
          processed: "length(body('Verify')?['value'])", loops: VOL, calls: `add(${VOL}, 1)`, actions: `add(${BASE + 3}, ${VOL})`, checksum: uc4Check,
          statusExpr: `if(less(length(body('Verify')?['value']), ${VOL}), 'Partial', 'Succeeded')`,
        }, after('Verify')),
      },
    }));
  }
  // PoC + lab: SharePoint works through a $batch serially (~100–170 ms per item) and stops after ~90 s,
  // so FAST uses batches of 250 and sends 4 of them in parallel.
  const BATCH = `min(max(int(coalesce(${pl('batchSize')}, 250)), 1), 500)`;
  add(benchmarkFlow(P, {
    name: 'UC4-FAST', scenario: 'UC4', variant: 'FAST', technique: 'SharePoint $batch (250 per changeset, 4 batches in parallel)',
    description: 'UC4 FAST: chunk() the rows into batches of 250 and send 4 $batch requests in parallel (SharePoint stops a batch after ~90 s).',
    apis: [SP], volumeDefault: 200, volumeMax: 5000,
    work: {
      ...crlf,
      Chunks: compose(`@chunk(range(0, ${VOL}), ${BATCH})`, after('CRLF')),
      Batch_loop: forEach("@outputs('Chunks')", {
        ...spBatch('B', "@items('Batch_loop')", createReq('OrderImport',
          `concat('{"__metadata":{"type":"${L.OrderImport.entityType}"},"Title":"', ${oiTitle('item()')}, '","Amount":', string(${data.amount('add(item(), 1)')}), ',"Region":"', ${data.region('item()')}, '","BatchTag":"', ${CORR}, '"}')`), {}),
      }, after('Chunks'), 4),
      Verify: uc4Verify(after('Batch_loop', ['Succeeded', 'Failed'])),
      Kpi: kpi({
        processed: "length(body('Verify')?['value'])", loops: "length(outputs('Chunks'))", calls: "add(length(outputs('Chunks')), 1)",
          statusExpr: `if(less(length(body('Verify')?['value']), ${VOL}), 'Partial', 'Succeeded')`,
        payload: '0', actions: `add(${BASE + 4}, mul(6, length(outputs('Chunks'))))`, checksum: uc4Check,
      }, after('Verify')),
    },
  }));

  // ================================================================= UC5 Dataverse bulk create
  const bad = (i) => `and(equals(${pl('injectBadItem')}, true), equals(${i}, 0))`;
  for (const [variant, conc] of [['SLOW', 1], ['MID', 50]]) {
    add(benchmarkFlow(P, {
      name: `UC5-${variant}`, scenario: 'UC5', variant,
      technique: conc === 1 ? 'Apply to each + Add a new row (concurrency 1)' : 'Apply to each + Add a new row (concurrency 50)',
      description: `UC5 ${variant}: one "Add a new row" per record${conc > 1 ? ', 50 in parallel' : ''}. With injectBadItem, row 0 fails and the other N-1 are still created.`,
      apis: [DV], volumeDefault: 500, volumeMax: 5000,
      work: {
        Create_loop: forEach(`@range(0, ${VOL})`, {
          Add_row: dvAction('CreateRecord', {
            entityName: `${P}_orderimports`,
            [`item/${P}_name`]: `@{${oiTitle("items('Create_loop')")}}`,
            [`item/${P}_externalid`]: `@{${oiTitle("items('Create_loop')")}}`,
            [`item/${P}_region`]: `@{${data.region("items('Create_loop')")}}`,
            [`item/${P}_amount`]: `@if(${bad("items('Create_loop')")}, 'not-a-number', ${data.amount("add(items('Create_loop'), 1)")})`,
            [`item/${P}_batchtag`]: `@{${CORR}}`,
          }, {}),
        }, {}, conc),
        Verify: verifyImport(after('Create_loop', ['Succeeded', 'Failed'])),
        Kpi: kpi({
          processed: verifyImportCount, loops: VOL, calls: `add(${VOL}, 1)`, actions: `add(${BASE + 3}, ${VOL})`, checksum: verifyImportChecksum,
          statusExpr: `if(less(${verifyImportCount}, ${VOL}), 'Partial', 'Succeeded')`,
        }, after('Verify')),
      },
    }));
  }
  const target = (i) => `concat('{"@odata.type":"Microsoft.Dynamics.CRM.${P}_orderimport","${P}_name":"', ${oiTitle(i)}, '","${P}_externalid":"', ${oiTitle(i)}, '","${P}_region":"', ${data.region(i)}, '","${P}_amount":', if(${bad(i)}, '"not-a-number"', string(${data.amount(`add(${i}, 1)`)})), ',"${P}_batchtag":"', ${CORR}, '"}')`;
  add(benchmarkFlow(P, {
    name: 'UC5-FAST', scenario: 'UC5', variant: 'FAST', technique: 'CreateMultiple (1,000 per call) via HTTP with Microsoft Entra ID',
    description: 'UC5 FAST: CreateMultiple with 1,000 targets per request. One request is one transaction: with injectBadItem the whole request rolls back.',
    apis: [HTTPDV, DV], volumeDefault: 500, volumeMax: 5000,
    work: {
      Chunks: compose(`@chunk(range(0, ${VOL}), 1000)`, {}),
      Bulk_loop: forEach("@outputs('Chunks')", {
        Targets: select("@items('Bulk_loop')", `@${target('item()')}`, {}),
        Create_multiple: dvHttp(cfg, 'POST', `${P}_orderimports/Microsoft.Dynamics.CRM.CreateMultiple`, "@{concat('{\"Targets\":[', join(body('Targets'), ','), ']}')}", after('Targets')),
      }, after('Chunks')),
      Verify: verifyImport(after('Bulk_loop', ['Succeeded', 'Failed'])),
      Kpi: kpi({
        processed: verifyImportCount, loops: "length(outputs('Chunks'))", calls: "add(length(outputs('Chunks')), 1)",
          statusExpr: `if(less(${verifyImportCount}, ${VOL}), 'Partial', 'Succeeded')`,
        actions: `add(${BASE + 3}, mul(2, length(outputs('Chunks'))))`, checksum: verifyImportChecksum,
      }, after('Verify')),
    },
  }));

  // ================================================================= UC6 Filter at source (SharePoint)
  const uc6Check = (arrExpr) => `concat(string(length(${arrExpr})), '|', ${money(xsum(arrExpr, 'amount'))}, '|', string(int(xpath(${xmlOf(arrExpr)}, 'count(/r/o[segment="Enterprise"])'))))`;
  add(benchmarkFlow(P, {
    name: 'UC6-SLOW', scenario: 'UC6', variant: 'SLOW', technique: 'Get items (all columns, paginated) → Filter array → Get item per row',
    description: 'UC6 SLOW: read every order in the window with all ~26 columns, filter in the flow (EMEA + Open), then fetch each customer with Get item.',
    apis: [SP], volumeDefault: 2000, volumeMax: 20000,
    vars: [{ name: 'Result', type: 'array', value: [] }],
    work: {
      Get_items: spAction(cfg, 'GetItems', { table: L.Orders.id, $filter: `ID le @{${VOL}}`, $top: 5000 }, {}, { runtimeConfiguration: { paginationPolicy: { minimumItemCount: 20000 } } }),
      Matches: filter("@body('Get_items')?['value']", "@and(equals(item()?['Region'], 'EMEA'), equals(item()?['Status'], 'Open'), lessOrEquals(int(item()?['ID']), " + VOL + '))', after('Get_items')),
      Lookup_loop: forEach("@body('Matches')", {
        Get_customer: spAction(cfg, 'GetItem', { table: L.Customers.id, id: "@coalesce(items('Lookup_loop')?['Customer']?['Id'], items('Lookup_loop')?['Customer#Id'])" }, {}),
        Enriched: compose({ orderId: "@items('Lookup_loop')?['Title']", amount: "@items('Lookup_loop')?['Amount']", customerName: "@body('Get_customer')?['CustomerName']", segment: "@body('Get_customer')?['Segment']" }, after('Get_customer')),
        Append_enriched: appendVar('Result', "@outputs('Enriched')", after('Enriched')),
      }, after('Matches')),
      Kpi: kpi({
        processed: "length(variables('Result'))", retrieved: "length(body('Get_items')?['value'])", loops: "length(body('Matches'))",
        calls: "add(div(add(length(body('Get_items')?['value']), 4999), 5000), length(body('Matches')))",
        payload: "length(string(body('Get_items')))", actions: `add(${BASE + 3}, mul(3, length(body('Matches'))))`,
        checksum: uc6Check("variables('Result')"),
      }, after('Lookup_loop', ['Succeeded', 'Failed'])),
    },
  }));
  add(benchmarkFlow(P, {
    name: 'UC6-FAST', scenario: 'UC6', variant: 'FAST', technique: 'REST $filter + $select + $expand (lookup) at source',
    description: 'UC6 FAST: one REST call filters (indexed Region first), selects 5 columns and expands the Customer lookup.',
    apis: [SP], volumeDefault: 2000, volumeMax: 20000,
    work: {
      Query: spHttp(cfg, 'GET', `${listUri('Orders')}/items?$select=Id,Title,Amount,Customer/CustomerName,Customer/Segment&$expand=Customer&$filter=Region eq 'EMEA' and Status eq 'Open' and Id le @{${VOL}}&$top=5000`, {}),
      Enriched: select("@body('Query')?['value']", { orderId: "@item()?['Title']", amount: "@item()?['Amount']", customerName: "@item()?['Customer']?['CustomerName']", segment: "@item()?['Customer']?['Segment']" }, after('Query')),
      Kpi: kpi({
        processed: "length(body('Enriched'))", retrieved: "length(body('Query')?['value'])", calls: '1',
        payload: "length(string(body('Query')))", actions: String(BASE + 3), checksum: uc6Check("body('Enriched')"),
      }, after('Enriched')),
    },
  }));

  // ================================================================= UC6-DV Aggregate at source (Dataverse)
  const lastOrder = `concat('O', ${pad(VOL, 5)})`;
  add(benchmarkFlow(P, {
    name: 'UC6DV-SLOW', scenario: 'UC6DV', variant: 'SLOW', technique: 'List rows (all columns, paginated) → Filter array → sum in the flow',
    description: 'UC6-DV SLOW: list every order row in the window with all columns, filter Open in the flow and total the amounts.',
    apis: [DV], volumeDefault: 5000, volumeMax: 20000,
    work: {
      List_rows: dvAction('ListRecords', { entityName: `${P}_orders`, $filter: `${P}_name le '@{${lastOrder}}'`, $orderby: `${P}_name` }, {}, { runtimeConfiguration: { paginationPolicy: { minimumItemCount: 20000 } } }),
      Open_rows: filter("@body('List_rows')?['value']", `@equals(item()?['${P}_status'], 'Open')`, after('List_rows')),
      Amounts: select("@body('Open_rows')", { a: `@item()?['${P}_amount']` }, after('Open_rows')),
      Kpi: kpi({
        processed: "length(body('Open_rows'))", retrieved: "length(body('List_rows')?['value'])", calls: "div(add(length(body('List_rows')?['value']), 4999), 5000)",
        payload: "length(string(body('List_rows')))", actions: String(BASE + 4),
        checksum: `concat(string(length(body('Amounts'))), '|', ${money(xsum("body('Amounts')", 'a'))})`,
      }, after('Amounts')),
    },
  }));
  add(benchmarkFlow(P, {
    name: 'UC6DV-FAST', scenario: 'UC6DV', variant: 'FAST', technique: 'FetchXML aggregate (count + sum) at source',
    description: 'UC6-DV FAST: one FetchXML aggregate query returns the count and the total; no rows leave Dataverse.',
    apis: [DV], volumeDefault: 5000, volumeMax: 20000,
    work: {
      Aggregate: dvAction('ListRecords', {
        entityName: `${P}_orders`,
        fetchXml: `<fetch aggregate="true"><entity name="${P}_order"><attribute name="${P}_orderid" alias="n" aggregate="count"/><attribute name="${P}_amount" alias="s" aggregate="sum"/><filter><condition attribute="${P}_status" operator="eq" value="Open"/><condition attribute="${P}_name" operator="le" value="@{${lastOrder}}"/></filter></entity></fetch>`,
      }, {}),
      Kpi: kpi({
        processed: "int(coalesce(first(body('Aggregate')?['value'])?['n'], 0))", retrieved: '1', calls: '1', payload: "length(string(body('Aggregate')))", actions: String(BASE + 2),
        checksum: "concat(string(coalesce(first(body('Aggregate')?['value'])?['n'], 0)), '|', " + money("coalesce(first(body('Aggregate')?['value'])?['s'], 0)") + ')',
      }, after('Aggregate')),
    },
  }));

  // ================================================================= UC7 Trigger conditions
  // DRIVER (launched from the app) modifies N items of UC7Items in one $batch; 1 in 10 becomes Approved.
  // UC7-SLOW and UC7-FAST are AUTOMATED flows on that list; each run that sees a CorrelationId logs a row
  // (label = the driver's correlationId), so the app can count runs per variant.
  const relevant = (i) => `equals(mod(${i}, 10), 0)`;
  add(benchmarkFlow(P, {
    name: 'UC7-DRIVER', scenario: 'UC7', variant: 'DRIVER', technique: 'Modifies N list items (1 in 10 relevant) for the trigger-condition demo',
    description: 'UC7 DRIVER: updates N items of UC7Items with one $batch (10 % to Approved). UC7-SLOW / UC7-FAST react to it.',
    apis: [SP], volumeDefault: 100, volumeMax: 100,
    work: {
      ...crlf,
      Get_items: spHttp(cfg, 'GET', `${listUri('UC7Items')}/items?$select=Id&$orderby=Id&$top=100`, after('CRLF')),
      ...spBatch('M', `@range(0, min(${VOL}, length(body('Get_items')?['value'])))`, mergeReq('UC7Items', "body('Get_items')?['value'][item()]?['Id']",
        `concat('{"__metadata":{"type":"${L.UC7Items.entityType}"},"Status":"', if(${relevant('item()')}, 'Approved', 'Pending'), '","Notes":"touched ', ${CORR}, ' #', string(item()), '","CorrelationId":"', ${CORR}, '"}')`), after('Get_items')),
      Kpi: kpi({
        processed: countIn("outputs('M_text')", 'HTTP/1.1 204'), loops: '0', calls: '2', actions: String(BASE + 7),
        checksum: `concat(string(${countIn("outputs('M_text')", 'HTTP/1.1 204')}), '|', string(div(add(min(${VOL}, length(body('Get_items')?['value'])), 9), 10)))`,
      }, after('M_text')),
    },
  }));
  const spTrigger = (conditions) => ({
    When_an_item_is_created_or_modified: {
      type: 'OpenApiConnection',
      recurrence: { interval: 1, frequency: 'Minute' },
      splitOn: "@triggerOutputs()?['body/value']",
      inputs: {
        host: { connectionName: SP, operationId: 'GetOnUpdatedItems', apiId: `/providers/Microsoft.PowerApps/apis/${SP}` },
        parameters: { dataset: site, table: L.UC7Items.id },
        authentication: "@parameters('$authentication')",
      },
      ...(conditions ? { conditions } : {}),
    },
  });
  const T = (f) => `triggerOutputs()?['body/${f}']`;
  const uc7Log = (variant, technique, statusExpr, runAfter) => dvAction('CreateRecord', {
    entityName: `${P}_benchmarkruns`,
    [`item/${P}_name`]: `@{concat(${T('CorrelationId')}, '-', ${T('ID')}, '-${variant}-', workflow()?['run']?['name'])}`,
    [`item/${P}_scenario`]: 'UC7', [`item/${P}_variant`]: variant, [`item/${P}_technique`]: technique,
    [`item/${P}_status`]: statusExpr, [`item/${P}_label`]: `@{${T('CorrelationId')}}`,
    [`item/${P}_starttime`]: '@utcNow()', [`item/${P}_endtime`]: '@utcNow()', [`item/${P}_volume`]: 1,
    [`item/${P}_runid`]: "@{workflow()?['run']?['name']}", [`item/${P}_flowid`]: "@{workflow()?['name']}",
    [`item/${P}_environmentid`]: "@{workflow()?['tags']?['environmentName']}",
    [`item/${P}_runurl`]: "@{concat('https://make.powerautomate.com/environments/', workflow()?['tags']?['environmentName'], '/flows/', workflow()?['name'], '/runs/', workflow()?['run']?['name'])}",
    [`item/${P}_details`]: `@{concat('{"itemId":', string(${T('ID')}), ',"status":"', ${T('Status')}, '"}')}`,
  }, runAfter);
  add({
    name: 'UC7-SLOW', description: 'UC7 SLOW (automated): runs on EVERY change to UC7Items, then checks Status in a Condition and terminates when irrelevant.',
    apis: [SP, DV], meta: { scenario: 'UC7', variant: 'SLOW', technique: 'Condition + Terminate inside the flow', app: false },
    trigger: spTrigger(),
    actions: {
      Has_correlation: cond({ not: { equals: [`@empty(${T('CorrelationId')})`, true] } }, {
        Log_run: uc7Log('SLOW', 'Condition + Terminate inside the flow', `@{if(equals(${T('Status')}, 'Approved'), 'Relevant', 'Ignored')}`, {}),
        Is_relevant: cond(and(eq(`@${T('Status')}`, 'Approved')), {
          Process: compose({ itemId: `@${T('ID')}`, processedAt: '@utcNow()' }, {}),
        }, {
          Not_relevant: { type: 'Terminate', inputs: { runStatus: 'Succeeded' }, runAfter: {} },
        }, after('Log_run')),
      }, {
        No_correlation: { type: 'Terminate', inputs: { runStatus: 'Succeeded' }, runAfter: {} },
      }, {}),
    },
  });
  add({
    name: 'UC7-FAST', description: 'UC7 FAST (automated): same trigger with trigger conditions; the flow only runs for relevant changes (Status = Approved).',
    apis: [SP, DV], meta: { scenario: 'UC7', variant: 'FAST', technique: 'Trigger conditions', app: false },
    trigger: spTrigger([
      { expression: `@equals(${T('Status')}, 'Approved')` },
      { expression: `@not(empty(${T('CorrelationId')}))` },
    ]),
    actions: {
      Log_run: uc7Log('FAST', 'Trigger conditions', 'Relevant', {}),
      Process: compose({ itemId: `@${T('ID')}`, processedAt: '@utcNow()' }, after('Log_run')),
    },
  });

  // ================================================================= UC8 Platform tuning traps
  const getItemsUc8 = (paginate) => spAction(cfg, 'GetItems', { table: L.Orders.id, $filter: `ID le @{${VOL}}`, ...(paginate ? { $top: 5000 } : {}) }, {},
    paginate ? { runtimeConfiguration: { paginationPolicy: { minimumItemCount: 5000 } } } : undefined);
  add(benchmarkFlow(P, {
    name: 'UC8-SLOW', scenario: 'UC8', variant: 'SLOW', technique: 'No pagination (silently 100 items) + concurrent loop updating a variable (race)',
    description: 'UC8 SLOW: Get items without pagination returns only 100 rows, and a concurrent loop (20) that reads + sets a variable loses updates. The checksum is WRONG on purpose.',
    apis: [SP], volumeDefault: 1000, volumeMax: 5000,
    vars: [{ name: 'Total', type: 'float', value: 0 }],
    work: {
      Get_items: getItemsUc8(false),
      Sum_loop: forEach("@body('Get_items')?['value']", {
        Current: compose("@variables('Total')", {}),
        Set_total: setVar('Total', "@add(outputs('Current'), float(items('Sum_loop')?['Amount']))", after('Current')),
      }, after('Get_items'), 20),
      Kpi: kpi({
        processed: "length(body('Get_items')?['value'])", retrieved: "length(body('Get_items')?['value'])", loops: "length(body('Get_items')?['value'])", calls: '1',
        payload: "length(string(body('Get_items')))", actions: `add(${BASE + 2}, mul(2, length(body('Get_items')?['value'])))`,
        checksum: `concat(string(length(body('Get_items')?['value'])), '|', ${money("variables('Total')")})`,
        extra: { expected: VOL, paginationTruncated: `less(length(body('Get_items')?['value']), ${VOL})` },
      }, after('Sum_loop')),
    },
  }));
  add(benchmarkFlow(P, {
    name: 'UC8-FAST', scenario: 'UC8', variant: 'FAST', technique: 'Pagination threshold + Select/xpath sum (no variables, no race)',
    description: 'UC8 FAST: pagination on (Top 5000) returns every row; the total comes from Select + xpath, so there is nothing to race on.',
    apis: [SP], volumeDefault: 1000, volumeMax: 5000,
    work: {
      Get_items: getItemsUc8(true),
      Amounts: select("@body('Get_items')?['value']", { a: "@float(item()?['Amount'])" }, after('Get_items')),
      Kpi: kpi({
        processed: "length(body('Amounts'))", retrieved: "length(body('Amounts'))", calls: '1', payload: "length(string(body('Get_items')))", actions: String(BASE + 3),
        checksum: `concat(string(length(body('Amounts'))), '|', ${money(xsum("body('Amounts')", 'a'))})`,
        extra: { expected: VOL },
      }, after('Amounts')),
    },
  }));

  // ================================================================= UC9 Excel
  if (cfg.uc9) {
    const X = cfg.uc9;
    const runScript = (variant, script, params, runAfter) => xlAction('RunScriptProdV2', {
      source: X.source, drive: X.driveId, file: X.files[variant], scriptSource: X.source, scriptDrive: X.scriptDriveId, scriptId: X.scripts[script], ...params,
    }, runAfter);
    const uc9Kpi = (variant, calls, loops, actions, runAfter) => kpi({
      processed: "int(first(split(string(body('Stats')?['result']), '|')))", loops, calls, actions,
      checksum: "string(body('Stats')?['result'])",
    }, runAfter);
    add(benchmarkFlow(P, {
      name: 'UC9-SLOW', scenario: 'UC9', variant: 'SLOW', technique: 'Apply to each + Add a row into a table',
      description: 'UC9 SLOW: one "Add a row into a table" call per row (Excel Online). The classic hours-to-seconds opener.',
      apis: [XL], volumeDefault: 200, volumeMax: 2000,
      work: {
        Reset: runScript('SLOW', 'Uc9Reset', {}, {}),
        Add_loop: forEach(`@range(0, ${VOL})`, {
          Add_row: xlAction('AddRowV2', {
            source: X.source, drive: X.driveId, file: X.files.SLOW, table: 'Orders',
            'item/OrderId': `@{${data.orderId("add(items('Add_loop'), 1)")}}`,
            'item/Region': `@{${data.region("add(items('Add_loop'), 1)")}}`,
            'item/Amount': `@{${data.amount("add(items('Add_loop'), 1)")}}`,
          }, {}),
        }, after('Reset')),
        Stats: runScript('SLOW', 'Uc9Stats', {}, after('Add_loop', ['Succeeded', 'Failed'])),
        Kpi: uc9Kpi('SLOW', `add(${VOL}, 2)`, VOL, `add(${BASE + 4}, ${VOL})`, after('Stats')),
      },
    }));
    add(benchmarkFlow(P, {
      name: 'UC9-FAST', scenario: 'UC9', variant: 'FAST', technique: 'Office Script setValues() in one call',
      description: 'UC9 FAST: Select builds all rows, one "Run script" call writes them with setValues().',
      apis: [XL], volumeDefault: 200, volumeMax: 2000,
      work: {
        Reset: runScript('FAST', 'Uc9Reset', {}, {}),
        Rows: select(`@range(1, ${VOL})`, `@createArray(${data.orderId('item()')}, ${data.region('item()')}, ${data.amount('item()')})`, {}),
        Write_rows: runScript('FAST', 'Uc9WriteRows', { 'ScriptParameters/rowsJson': "@{string(body('Rows'))}" }, { Reset: ['Succeeded'], Rows: ['Succeeded'] }),
        Stats: runScript('FAST', 'Uc9Stats', {}, after('Write_rows', ['Succeeded', 'Failed'])),
        Kpi: uc9Kpi('FAST', '3', '0', String(BASE + 5), after('Stats')),
      },
    }));
  }

  // ================================================================= UC10 Fan-out
  const work10 = (i) => `mod(mul(${i}, 7), 13)`;
  add(benchmarkFlow(P, {
    name: 'UC10-SLOW', scenario: 'UC10', variant: 'SLOW', technique: 'One flow, one sequential loop over all records',
    description: 'UC10 SLOW: a single sequential Apply to each over N records (2 actions per record).',
    apis: [], volumeDefault: 1000, volumeMax: 20000,
    vars: [{ name: 'Total', type: 'integer', value: 0 }],
    work: {
      Loop: forEach(`@range(0, ${VOL})`, {
        Value: compose(`@${work10("items('Loop')")}`, {}),
        Add_value: incVar('Total', "@outputs('Value')", after('Value')),
      }, {}),
      Kpi: kpi({ processed: VOL, loops: VOL, calls: '0', actions: `add(${BASE + 1}, mul(2, ${VOL}))`, checksum: `concat(string(${VOL}), '|', string(variables('Total')))` }, after('Loop')),
    },
  }));
  add({
    name: 'UC10-CHILD', description: 'UC10 child flow: processes one chunk (start, count) sequentially and returns its partial sum. Called by UC10-FAST.',
    apis: [], meta: { scenario: 'UC10', variant: 'CHILD', technique: 'Child flow per chunk', app: false },
    trigger: powerAppsTrigger,
    actions: {
      Payload: compose("@json(coalesce(triggerBody()?['text'], '{}'))", {}),
      Init_Total: { type: 'InitializeVariable', inputs: { variables: [{ name: 'Total', type: 'integer', value: 0 }] }, runAfter: after('Payload') },
      Loop: forEach("@range(int(outputs('Payload')?['start']), int(outputs('Payload')?['count']))", {
        Value: compose(`@${work10("items('Loop')")}`, {}),
        Add_value: incVar('Total', "@outputs('Value')", after('Value')),
      }, after('Init_Total')),
      Respond: {
        type: 'Response', kind: 'PowerApp',
        inputs: { statusCode: 200, body: { sum: "@variables('Total')", count: "@int(outputs('Payload')?['count'])" }, schema: { type: 'object', properties: { sum: { title: 'sum', type: 'number', 'x-ms-dynamically-added': true }, count: { title: 'count', type: 'number', 'x-ms-dynamically-added': true } } } },
        runAfter: after('Loop'),
      },
    },
  });
  if (ids['UC10-CHILD']) {
    add(benchmarkFlow(P, {
      name: 'UC10-FAST', scenario: 'UC10', variant: 'FAST', technique: 'chunk() + Apply to each (concurrency 10) → child flow per chunk',
      description: 'UC10 FAST: split the records into 10 chunks and process them in parallel child flows.',
      apis: [], volumeDefault: 1000, volumeMax: 20000,
      work: {
        Size: compose(`@max(div(add(${VOL}, 9), 10), 1)`, {}),
        Chunks: compose(`@chunk(range(0, ${VOL}), outputs('Size'))`, after('Size')),
        Fan_out: forEach("@outputs('Chunks')", {
          Run_child: {
            type: 'Workflow',
            inputs: { host: { workflowReferenceName: ids['UC10-CHILD'] }, body: { text: "@{concat('{\"start\":', string(first(items('Fan_out'))), ',\"count\":', string(length(items('Fan_out'))), '}')}" } },
            runAfter: {},
          },
        }, after('Chunks'), 10),
        // result() of a loop returns one entry per action; its `outputs` array holds every iteration.
        Child_results: filter("@result('Fan_out')", "@equals(item()?['name'], 'Run_child')", after('Fan_out', ['Succeeded', 'Failed'])),
        Partials: select("@coalesce(first(body('Child_results'))?['outputs'], json('[]'))", { s: "@item()?['outputs']?['body']?['sum']", c: "@item()?['outputs']?['body']?['count']" }, after('Child_results')),
        Kpi: kpi({
          processed: `int(${xsum("body('Partials')", 'c')})`, loops: "length(outputs('Chunks'))", calls: "length(outputs('Chunks'))",
          actions: `add(add(${BASE + 4}, length(outputs('Chunks'))), mul(2, ${VOL}))`,
          checksum: `concat(string(int(${xsum("body('Partials')", 'c')})), '|', string(int(${xsum("body('Partials')", 's')})))`,
        }, after('Partials')),
      },
    }));
  }

  // ================================================================= SEED-SP
  const custJson = (n) => `concat('{"__metadata":{"type":"${L.Customers.entityType}"},"Title":"', ${data.custCode(n)}, '","CustomerName":"Customer ', ${pad(n, 4)}, '","Segment":"', ${data.segment(n)}, '","Country":"', ${data.country(n)}, '"}')`;
  const noise = Array.from({ length: 20 }, (_, k) => `',"Noise${String(k + 1).padStart(2, '0')}":"Noise ${k + 1} lorem ipsum dolor sit amet #', string(item()), '"'`).join(', ');
  const orderJson = `concat('{"__metadata":{"type":"${L.Orders.entityType}"},"Title":"', ${data.orderId('item()')}, '","CustCode":"', ${data.custCode(data.custNo('item()'))}, '","CustomerId":', string(${data.custNo('item()')}), ',"Region":"', ${data.region('item()')}, '","Status":"', ${data.status('item()')}, '","Amount":', string(${data.amount('item()')}), ',"OrderDate":"', ${data.orderDate('item()')}, 'T00:00:00Z"', ${noise}, '}')`;
  const uc7Json = `concat('{"__metadata":{"type":"${L.UC7Items.entityType}"},"Title":"Item ', ${pad('item()', 3)}, '","Status":"Pending","Notes":"","CorrelationId":""}')`;
  const itemCount = (list) => spHttp(cfg, 'GET', `${listUri(list)}?$select=ItemCount`, {});
  add(benchmarkFlow(P, {
    name: 'SEED-SP', scenario: 'ADMIN', variant: 'SEED-SP', technique: 'Seed SharePoint lists with $batch (Customers 2,000, Orders N, UC7Items 100)',
    description: 'Seeds the SharePoint demo lists that are empty: Customers (2,000), Orders (volume, default 20,000) and UC7Items (100). Lists that already have items are skipped.',
    apis: [SP], volumeDefault: 20000, volumeMax: 20000,
    work: {
      ...crlf,
      Count_customers: itemCount('Customers'),
      Count_orders: itemCount('Orders'),
      Count_uc7: itemCount('UC7Items'),
      Seed_customers: cond({ and: [{ equals: ["@body('Count_customers')?['ItemCount']", 0] }] }, {
        Customer_loop: forEach('@chunk(range(1, 2000), 400)', spBatch('C', "@items('Customer_loop')", createReq('Customers', custJson('item()')), {}), {}),
      }, {}, { CRLF: ['Succeeded'], Count_customers: ['Succeeded'] }),
      Seed_orders: cond({ and: [{ equals: ["@body('Count_orders')?['ItemCount']", 0] }] }, {
        Order_loop: forEach(`@chunk(range(1, ${VOL}), 250)`, spBatch('O', "@items('Order_loop')", createReq('Orders', orderJson), {}), {}, 4),
      }, {}, { Seed_customers: ['Succeeded'], Count_orders: ['Succeeded'] }),
      Seed_uc7: cond({ and: [{ equals: ["@body('Count_uc7')?['ItemCount']", 0] }] }, {
        ...spBatch('U', '@range(1, 100)', createReq('UC7Items', uc7Json), {}),
      }, {}, { CRLF: ['Succeeded'], Count_uc7: ['Succeeded'] }),
      Final_customers: itemCount('Customers'), Final_orders: itemCount('Orders'), Final_uc7: itemCount('UC7Items'),
      Kpi: kpi({
        processed: "body('Final_orders')?['ItemCount']", calls: '9', actions: String(BASE + 20),
        checksum: "concat(string(body('Final_customers')?['ItemCount']), '|', string(body('Final_orders')?['ItemCount']), '|', string(body('Final_uc7')?['ItemCount']))",
      }, { Final_customers: ['Succeeded'], Final_orders: ['Succeeded'], Final_uc7: ['Succeeded'] }),
    },
  }));
  // Final_* must wait for the seeding branches
  const seedSp = flows[flows.length - 1].actions.Work.actions;
  for (const k of ['Final_customers', 'Final_orders', 'Final_uc7']) seedSp[k].runAfter = { Seed_customers: ANY, Seed_orders: ANY, Seed_uc7: ANY };

  // ================================================================= SEED-DV
  const dvOrder = (i) => `concat('{"@odata.type":"Microsoft.Dynamics.CRM.${P}_order","${P}_name":"', ${data.orderId(i)}, '","${P}_customerid":"', ${data.custCode(data.custNo(i))}, '","${P}_region":"', ${data.region(i)}, '","${P}_status":"', ${data.status(i)}, '","${P}_amount":', string(${data.amount(i)}), ',"${P}_orderdate":"', ${data.orderDate(i)}, '"}')`;
  const dvCount = (name) => dvAction('ListRecords', { entityName: `${P}_orders`, fetchXml: `<fetch aggregate="true"><entity name="${P}_order"><attribute name="${P}_orderid" alias="n" aggregate="count"/></entity></fetch>` }, name === 'Count_rows' ? {} : { Seed: ANY });
  add(benchmarkFlow(P, {
    name: 'SEED-DV', scenario: 'ADMIN', variant: 'SEED-DV', technique: 'Seed pa_order with CreateMultiple (1,000 per call, 4 in parallel)',
    description: 'Seeds the Dataverse pa_order table (volume rows, default 20,000) with CreateMultiple when it is empty.',
    apis: [HTTPDV, DV], volumeDefault: 20000, volumeMax: 20000,
    work: {
      Count_rows: dvCount('Count_rows'),
      Seed: cond({ and: [{ equals: ["@int(coalesce(first(body('Count_rows')?['value'])?['n'], 0))", 0] }] }, {
        Seed_loop: forEach(`@chunk(range(1, ${VOL}), 1000)`, {
          Targets: select("@items('Seed_loop')", `@${dvOrder('item()')}`, {}),
          Create_multiple: dvHttp(cfg, 'POST', `${P}_orders/Microsoft.Dynamics.CRM.CreateMultiple`, "@{concat('{\"Targets\":[', join(body('Targets'), ','), ']}')}", after('Targets')),
        }, {}, 4),
      }, {}, after('Count_rows')),
      Final_count: dvCount('Final_count'),
      Kpi: kpi({
        processed: "int(coalesce(first(body('Final_count')?['value'])?['n'], 0))", calls: `add(2, div(add(${VOL}, 999), 1000))`, actions: String(BASE + 6),
        checksum: "string(coalesce(first(body('Final_count')?['value'])?['n'], 0))",
      }, after('Final_count')),
    },
  }));

  // ================================================================= RESET
  const tgt = (t) => ({ or: [{ equals: [`@coalesce(${pl('target')}, 'all')`, 'all'] }, { equals: [`@coalesce(${pl('target')}, 'all')`, t] }] });
  const resetWork = {
    ...crlf,
    Reset_uc4: cond(tgt('uc4'), {
      Delete_until_empty: {
        type: 'Until',
        expression: "@equals(length(body('Get_ids')?['value']), 0)",
        limit: { count: 60, timeout: 'PT1H' },
        actions: {
          Get_ids: spHttp(cfg, 'GET', `${listUri('OrderImport')}/items?$select=Id&$top=300`, {}),
          Has_ids: cond({ and: [{ greater: ["@length(body('Get_ids')?['value'])", 0] }] },
            spBatch('D', "@body('Get_ids')?['value']", deleteReq('OrderImport', "item()?['Id']"), {}), {}, after('Get_ids')),
        },
        runAfter: {},
      },
    }, {}, after('CRLF')),
    Reset_uc5: cond(tgt('uc5'), {
      Bulk_delete: dvHttp(cfg, 'POST', 'BulkDelete', `@{concat('{"QuerySet":[{"EntityName":"${P}_orderimport","ColumnSet":{"AllColumns":false,"Columns":[]},"Criteria":{"FilterOperator":"And","Conditions":[],"Filters":[]}}],"JobName":"FlowOptLab reset ${P}_orderimport","SendEmailNotification":false,"ToRecipients":[],"CCRecipients":[],"RecurrencePattern":"","StartDateTime":"', utcNow(), '","RunNow":true}')}`, {}),
    }, {}, {}),
    Reset_uc7: cond(tgt('uc7'), {
      Get_uc7: spHttp(cfg, 'GET', `${listUri('UC7Items')}/items?$select=Id&$top=100`, {}),
      ...spBatch('R', "@body('Get_uc7')?['value']", mergeReq('UC7Items', "item()?['Id']", `'{"__metadata":{"type":"${L.UC7Items.entityType}"},"Status":"Pending","Notes":"","CorrelationId":""}'`), after('Get_uc7')),
    }, {}, after('CRLF')),
    ...(cfg.uc9 ? {
      Reset_uc9: cond(tgt('uc9'), {
        Reset_slow: xlAction('RunScriptProdV2', { source: cfg.uc9.source, drive: cfg.uc9.driveId, file: cfg.uc9.files.SLOW, scriptSource: cfg.uc9.source, scriptDrive: cfg.uc9.scriptDriveId, scriptId: cfg.uc9.scripts.Uc9Reset }, {}),
        Reset_fast: xlAction('RunScriptProdV2', { source: cfg.uc9.source, drive: cfg.uc9.driveId, file: cfg.uc9.files.FAST, scriptSource: cfg.uc9.source, scriptDrive: cfg.uc9.scriptDriveId, scriptId: cfg.uc9.scripts.Uc9Reset }, {}),
      }, {}, {}),
    } : {}),
  };
  const resetDone = { Reset_uc4: ANY, Reset_uc5: ANY, Reset_uc7: ANY, ...(cfg.uc9 ? { Reset_uc9: ANY } : {}) };
  resetWork.Remaining_uc4 = spHttp(cfg, 'GET', `${listUri('OrderImport')}?$select=ItemCount`, resetDone);
  resetWork.Kpi = kpi({
    processed: '0', calls: '0', actions: String(BASE + 10),
    checksum: `concat('target=', coalesce(${pl('target')}, 'all'), '|OrderImport=', string(body('Remaining_uc4')?['ItemCount']), '|uc5=', if(equals(actions('Reset_uc5')?['status'], 'Succeeded'), 'BulkDelete job submitted', 'skipped'))`,
  }, after('Remaining_uc4'));
  add(benchmarkFlow(P, {
    name: 'RESET', scenario: 'ADMIN', variant: 'RESET', technique: 'Reset demo targets (payload.target = uc4 | uc5 | uc7 | uc9 | all)',
    description: 'Empties OrderImport (uc4), submits a BulkDelete for pa_orderimport (uc5), sets UC7Items back to Pending (uc7) and clears the UC9 workbooks (uc9).',
    apis: cfg.uc9 ? [SP, HTTPDV, XL] : [SP, HTTPDV], volumeDefault: 1, volumeMax: 1, work: resetWork,
  }));

  // ================================================================= LAB-TestRunner (not in the app)
  // HTTP trigger → calls one lab flow as a child flow. Lets scripts run the lab without the Code App
  // (the Code App calls flows through the Power Apps player only). Responds 202 at once; results land
  // in pa_benchmarkrun like an app run.
  // runnable: names of flows that are turned on — a child flow that was never turned on blocks activation.
  const appFlows = flows.filter((f) => f.meta.app);
  if (appFlows.every((f) => ids[f.name])) {
    const targets = appFlows.filter((f) => !runnable || runnable.includes(f.name));
    const runner = {
      Accepted: { type: 'Response', kind: 'Http', inputs: { statusCode: 202, body: { flow: "@triggerBody()?['flow']", accepted: '@utcNow()' } }, runAfter: {} },
    };
    for (const f of targets) {
      const key = f.name.replace(/[^A-Za-z0-9]/g, '_');
      runner[`If_${key}`] = cond(and(eq("@triggerBody()?['flow']", f.name)), {
        [`Run_${key}`]: { type: 'Workflow', inputs: { host: { workflowReferenceName: ids[f.name] }, body: { text: "@{string(triggerBody()?['payload'])}" } }, runAfter: {} },
      }, {}, after('Accepted'));
    }
    add({
      name: 'LAB-TestRunner', description: 'Test harness: POST {"flow":"UC1-FAST","payload":{...}} to its HTTP trigger to run a lab flow as a child flow (used by lab/scripts/30-run.mjs).',
      apis: [], meta: { scenario: 'ADMIN', variant: 'RUNNER', app: false },
      trigger: { manual: { type: 'Request', kind: 'Http', inputs: { schema: { type: 'object', properties: { flow: { type: 'string' }, payload: { type: 'object' } } } } } },
      actions: runner,
    });
  }

  return flows;
}

/** Wraps a built flow as workflow.clientdata. */
export function clientData(flow, refForApi) {
  return {
    properties: {
      connectionReferences: Object.fromEntries(flow.apis.map((api) => [api, {
        runtimeSource: 'embedded',
        connection: { connectionReferenceLogicalName: refForApi[api] },
        api: { name: api },
      }])),
      definition: {
        $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        contentVersion: '1.0.0.0',
        parameters: { $connections: { defaultValue: {}, type: 'Object' }, $authentication: { defaultValue: {}, type: 'SecureObject' } },
        triggers: flow.trigger,
        actions: flow.actions,
      },
      templateName: null,
    },
    schemaVersion: '1.0.0.0',
  };
}
