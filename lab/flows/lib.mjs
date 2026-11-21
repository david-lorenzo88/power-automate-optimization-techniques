// Building blocks for the lab flows (Logic Apps workflow definition language).
//
// Every app-launched flow follows the contract in docs/03-code-app-benchmark-harness.md,
// adjusted with the PoC results (poc/RESULTS.md):
//   Power Apps V2 trigger with ONE text input `payload` (JSON string)
//   → Start / Payload / Volume / Links
//   → Create_run_row (status Running) → Respond_to_app (respond EARLY: Run() > 120 s fails, #5)
//   → [Initialize variables]  (must be top level)
//   → Work (scope) → Kpi compose (last action in Work)
//   → Work_errors → Update_run_row (Succeeded/Failed + KPIs + checksum)
// Run links use the /flows/{id}/runs/{run} and /flows/{id}/details forms (the
// /solutions/~preferred/ and /edit forms were broken in the PoC, #6).

export const DV = 'shared_commondataserviceforapps';
export const SP = 'shared_sharepointonline';
export const HTTPDV = 'shared_webcontents';
export const XL = 'shared_excelonlinebusiness';

export const REF_FOR_API = (P) => ({
  [DV]: `${P}_dataverse`,
  [SP]: `${P}_sharepoint`,
  [HTTPDV]: `${P}_httpdataverse`,
  [XL]: `${P}_excel`,
});

/** WDL string literal: 'it''s' */
export const q = (s) => `'${String(s).replace(/'/g, "''")}'`;
export const pl = (k) => `outputs('Payload')?['${k}']`;
export const VOL = "outputs('Volume')";
export const CORR = "outputs('Payload')?['correlationId']";
/** Zero-padded number: pad(expr, 5) → '00042' */
export const pad = (expr, n) => `substring(string(add(${10 ** n}, ${expr})), 1, ${n})`;

// ------------------------------------------------------------------ connector actions
const conn = (api, operationId, parameters, runAfter, extra = {}) => ({
  type: 'OpenApiConnection',
  inputs: {
    host: { connectionName: api, operationId, apiId: `/providers/Microsoft.PowerApps/apis/${api}` },
    parameters,
    authentication: "@parameters('$authentication')",
  },
  runAfter,
  ...extra,
});

export const dvAction = (operationId, parameters, runAfter, extra) => conn(DV, operationId, parameters, runAfter, extra);
export const xlAction = (operationId, parameters, runAfter, extra) => conn(XL, operationId, parameters, runAfter, extra);

/** Dataverse Web API call through "HTTP with Microsoft Entra ID (preauthorized)". */
export const dvHttp = (cfg, method, path, body, runAfter, extra) => conn(HTTPDV, 'InvokeHttp', {
  'request/method': method,
  'request/url': `${cfg.dataverseUrl}/api/data/v9.2/${path}`,
  'request/headers': { 'Content-Type': 'application/json; charset=utf-8', Accept: 'application/json', 'OData-Version': '4.0', 'OData-MaxVersion': '4.0' },
  ...(body !== undefined ? { 'request/body': body } : {}),
}, runAfter, extra);

/** "Send an HTTP request to SharePoint" on the lab site. */
export const spHttp = (cfg, method, uri, runAfter, { headers, body, extra } = {}) => conn(SP, 'HttpRequest', {
  dataset: cfg.siteUrl,
  'parameters/method': method,
  'parameters/uri': uri,
  'parameters/headers': headers || { Accept: 'application/json;odata=nometadata' },
  ...(body !== undefined ? { 'parameters/body': body } : {}),
}, runAfter, extra);

export const spAction = (cfg, operationId, parameters, runAfter, extra) => conn(SP, operationId, { dataset: cfg.siteUrl, ...parameters }, runAfter, extra);

// ------------------------------------------------------------------ plain actions
export const compose = (inputs, runAfter) => ({ type: 'Compose', inputs, runAfter });
export const select = (from, sel, runAfter) => ({ type: 'Select', inputs: { from, select: sel }, runAfter });
export const filter = (from, where, runAfter) => ({ type: 'Query', inputs: { from, where }, runAfter });
export const initVar = (name, type, value, runAfter) => ({ type: 'InitializeVariable', inputs: { variables: [{ name, type, ...(value !== undefined ? { value } : {}) }] }, runAfter });
export const setVar = (name, value, runAfter) => ({ type: 'SetVariable', inputs: { name, value }, runAfter });
export const incVar = (name, value, runAfter) => ({ type: 'IncrementVariable', inputs: { name, value }, runAfter });
export const appendVar = (name, value, runAfter) => ({ type: 'AppendToArrayVariable', inputs: { name, value }, runAfter });
export const scope = (actions, runAfter) => ({ type: 'Scope', actions, runAfter });
export const forEach = (items, actions, runAfter, concurrency = 1) => ({
  type: 'Foreach', foreach: items, actions, runAfter,
  runtimeConfiguration: { concurrency: { repetitions: concurrency } },
});
export const cond = (expression, yes, no, runAfter) => ({ type: 'If', expression, actions: yes, else: { actions: no || {} }, runAfter });
export const after = (name, statuses = ['Succeeded']) => ({ [name]: statuses });
export const ANY = ['Succeeded', 'Failed', 'Skipped', 'TimedOut'];

/** CRLF as a Compose so multipart bodies can be built with concat(). */
export const CRLF = "outputs('CRLF')";

/** xml() wrapper around an array of objects so xpath can sum/count it: /r/o/<field> */
export const xmlOf = (arrayExpr) => `xml(json(concat('{"r":{"o":', string(${arrayExpr}), '}}')))`;
export const xsum = (arrayExpr, field) => `xpath(${xmlOf(arrayExpr)}, 'sum(/r/o/${field})')`;
export const money = (expr) => `formatNumber(float(${expr}), '0.00', 'en-US')`;

// ------------------------------------------------------------------ deterministic demo data
// The same formulas are used by the seed flows and by flows that generate data in memory,
// so SLOW and FAST always see the same records.
export const REGIONS = ['EMEA', 'AMER', 'APAC', 'LATAM', 'MEA'];
export const SEGMENTS = ['Enterprise', 'SMB', 'Public'];
export const COUNTRIES = ['Spain', 'Latvia', 'Germany', 'USA', 'Brazil', 'Japan'];
const arr = (xs) => `createArray(${xs.map(q).join(', ')})`;
export const data = {
  region: (i) => `${arr(REGIONS)}[mod(${i}, 5)]`,
  // Open 15 %, Closed 50 %, Pending 25 %, Approved 10 %
  status: (i) => `if(less(mod(mul(${i}, 7), 20), 3), 'Open', if(less(mod(mul(${i}, 7), 20), 13), 'Closed', if(less(mod(mul(${i}, 7), 20), 18), 'Pending', 'Approved')))`,
  amount: (i) => `add(mod(mul(${i}, 37), 1000), div(float(mod(${i}, 100)), 100))`,
  orderId: (i) => `concat('O', ${pad(i, 5)})`,
  custNo: (i) => `add(mod(sub(${i}, 1), 2000), 1)`,
  custCode: (n) => `concat('C', ${pad(n, 4)})`,
  orderDate: (i) => `addDays('2026-01-01', mod(${i}, 300), 'yyyy-MM-dd')`,
  segment: (n) => `${arr(SEGMENTS)}[mod(${n}, 3)]`,
  country: (n) => `${arr(COUNTRIES)}[mod(${n}, 6)]`,
};

// ------------------------------------------------------------------ the contract
const env = "workflow()?['tags']?['environmentName']";
const flowName = "workflow()?['name']";
const runName = "workflow()?['run']?['name']";

export const powerAppsTrigger = {
  manual: {
    type: 'Request',
    kind: 'PowerAppV2',
    inputs: {
      schema: {
        type: 'object',
        properties: {
          text: { title: 'payload', type: 'string', description: 'JSON payload from the Code App', 'x-ms-dynamically-added': true, 'x-ms-content-hint': 'TEXT' },
        },
        required: ['text'],
      },
    },
  },
};

export function header({ volumeDefault = 500, volumeMax = 5000 } = {}) {
  return {
    Start: compose('@utcNow()', {}),
    Payload: compose("@json(coalesce(triggerBody()?['text'], '{}'))", after('Start')),
    Volume: compose(`@min(max(int(coalesce(${pl('volume')}, ${volumeDefault})), 1), ${volumeMax})`, after('Payload')),
    Links: compose({
      environmentId: `@{${env}}`,
      flowId: `@{${flowName}}`,
      runId: `@{${runName}}`,
      runUrl: `@{concat('https://make.powerautomate.com/environments/', ${env}, '/flows/', ${flowName}, '/runs/', ${runName})}`,
      detailsUrl: `@{concat('https://make.powerautomate.com/environments/', ${env}, '/flows/', ${flowName}, '/details')}`,
    }, after('Volume')),
  };
}

export function createRunRow(P, { scenario, variant, technique }) {
  return dvAction('CreateRecord', {
    entityName: `${P}_benchmarkruns`,
    [`item/${P}_name`]: `@{coalesce(${CORR}, guid())}`,
    [`item/${P}_scenario`]: scenario,
    [`item/${P}_variant`]: variant,
    [`item/${P}_technique`]: technique,
    [`item/${P}_status`]: 'Running',
    [`item/${P}_label`]: `@{${pl('label')}}`,
    [`item/${P}_volume`]: `@${VOL}`,
    [`item/${P}_starttime`]: "@outputs('Start')",
    [`item/${P}_runid`]: "@outputs('Links')?['runId']",
    [`item/${P}_flowid`]: "@outputs('Links')?['flowId']",
    [`item/${P}_environmentid`]: "@outputs('Links')?['environmentId']",
    [`item/${P}_runurl`]: "@outputs('Links')?['runUrl']",
    [`item/${P}_editurl`]: "@outputs('Links')?['detailsUrl']",
  }, after('Links'));
}

export const respondEarly = (runAfter) => ({
  type: 'Response',
  kind: 'PowerApp',
  inputs: {
    statusCode: 200,
    body: { runid: "@{outputs('Links')?['runId']}", runurl: "@{outputs('Links')?['runUrl']}", rowid: `@{outputs('Create_run_row')?['body/pa_benchmarkrunid']}` },
    schema: {
      type: 'object',
      properties: {
        runid: { title: 'runId', type: 'string', 'x-ms-dynamically-added': true },
        runurl: { title: 'runUrl', type: 'string', 'x-ms-dynamically-added': true },
        rowid: { title: 'rowId', type: 'string', 'x-ms-dynamically-added': true },
      },
    },
  },
  runAfter,
});

/**
 * Kpi compose: the last action inside Work. Every field is an expression string (no leading @).
 * statusExpr (optional): overrides Succeeded/Failed (e.g. 'Partial').
 */
export function kpi({ processed, retrieved = '0', loops = '0', calls = '0', actions = '0', payload = '0', checksum, extra = {}, statusExpr }, runAfter) {
  return compose({
    recordsProcessed: `@${processed}`,
    recordsRetrieved: `@${retrieved}`,
    loopIterations: `@${loops}`,
    connectorCalls: `@${calls}`,
    actionsEstimated: `@${actions}`,
    payloadBytes: `@${payload}`,
    checksum: `@{${checksum}}`,
    ...(statusExpr ? { status: `@{${statusExpr}}` } : {}),
    ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, `@${v}`])),
  }, runAfter);
}

/** Work_errors + Update_run_row. Must run after the Work scope. */
export function footer(P) {
  const k = (f) => `outputs('Kpi')?['${f}']`;
  const errs = "body('Work_errors')";
  return {
    Work_errors: {
      type: 'Query',
      inputs: { from: "@result('Work')", where: "@equals(item()?['status'], 'Failed')" },
      runAfter: after('Work', ANY),
    },
    Update_run_row: dvAction('UpdateRecord', {
      entityName: `${P}_benchmarkruns`,
      recordId: `@outputs('Create_run_row')?['body/${P}_benchmarkrunid']`,
      [`item/${P}_status`]: `@{coalesce(${k('status')}, if(equals(actions('Work')?['status'], 'Succeeded'), 'Succeeded', 'Failed'))}`,
      [`item/${P}_endtime`]: '@utcNow()',
      [`item/${P}_durationms`]: "@div(sub(ticks(utcNow()), ticks(outputs('Start'))), 10000)",
      [`item/${P}_recordsprocessed`]: `@${k('recordsProcessed')}`,
      [`item/${P}_recordsretrieved`]: `@${k('recordsRetrieved')}`,
      [`item/${P}_loopiterations`]: `@${k('loopIterations')}`,
      [`item/${P}_connectorcalls`]: `@${k('connectorCalls')}`,
      [`item/${P}_actionsestimated`]: `@${k('actionsEstimated')}`,
      [`item/${P}_payloadbytes`]: `@${k('payloadBytes')}`,
      [`item/${P}_resultchecksum`]: `@{${k('checksum')}}`,
      [`item/${P}_details`]: "@{string(json(concat('{\"kpi\":', string(coalesce(outputs('Kpi'), json('{}'))), ',\"payload\":', string(outputs('Payload')), ',\"links\":', string(outputs('Links')), '}')))}",
      [`item/${P}_error`]: `@{if(empty(${errs}), '', substring(string(${errs}), 0, min(length(string(${errs})), 4000)))}`,
    }, after('Work_errors', ANY)),
    // If Kpi never ran (Work failed early), the update above can't evaluate its inputs:
    // still close the row so the app stops polling.
    Update_run_row_fallback: dvAction('UpdateRecord', {
      entityName: `${P}_benchmarkruns`,
      recordId: `@outputs('Create_run_row')?['body/${P}_benchmarkrunid']`,
      [`item/${P}_status`]: 'Failed',
      [`item/${P}_endtime`]: '@utcNow()',
      [`item/${P}_durationms`]: "@div(sub(ticks(utcNow()), ticks(outputs('Start'))), 10000)",
      [`item/${P}_error`]: `@{if(empty(${errs}), 'Work failed', substring(string(${errs}), 0, min(length(string(${errs})), 4000)))}`,
    }, after('Update_run_row', ['Failed', 'Skipped', 'TimedOut'])),
  };
}

/**
 * Assemble an app-launched flow.
 * vars: [{ name, type, value }] initialised (top level) between Respond_to_app and Work.
 */
export function benchmarkFlow(P, { name, description, scenario, variant, technique, apis, volumeDefault, volumeMax, vars = [], work }) {
  const actions = {
    ...header({ volumeDefault, volumeMax }),
    Create_run_row: createRunRow(P, { scenario, variant, technique }),
    Respond_to_app: respondEarly(after('Create_run_row')),
  };
  let prev = 'Respond_to_app';
  vars.forEach((v, i) => {
    const n = `Init_${v.name}`;
    actions[n] = initVar(v.name, v.type, v.value, after(prev));
    prev = n;
  });
  actions.Work = scope(work, after(prev));
  Object.assign(actions, footer(P));
  return { name, description, apis: [...new Set([DV, ...apis])], meta: { scenario, variant, technique, app: true }, trigger: powerAppsTrigger, actions };
}
