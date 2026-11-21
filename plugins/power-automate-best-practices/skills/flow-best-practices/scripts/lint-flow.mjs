#!/usr/bin/env node
// Power Automate flow best-practice linter (performance, throughput, correctness traps).
// Zero dependencies, Node 18+. Works on:
//   - a raw workflow definition ({ triggers, actions, ... })
//   - Dataverse clientdata ({ properties: { definition } }) or its JSON string
//   - FlowAgent / Flow API get_flow output (anything that contains a definition)
//
// Usage:
//   node lint-flow.mjs <flow.json | -> [--format md|json] [--fail-on error|warning|none]
// Exit codes: 0 = no finding at/above --fail-on (default: error), 1 = findings at/above it, 2 = bad input.

import { readFileSync, realpathSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

// ------------------------------------------------------------------ rules
export const RULES = {
  BP001: { severity: 'warning', title: 'Loop used for data shaping', fix: 'Replace the loop with Filter array + Select (advanced mode: addProperty/setProperty), union() for distinct values, or xml()+xpath / length() / max() for aggregates. One action instead of N × body actions.' },
  BP002: { severity: 'warning', title: 'Nested loop (O(n × m))', fix: 'Build a lookup object once (Select key/value + json(concat(\'{\', join(...), \'}\'))) and read it with outputs(\'Dict\')?[key] inside a single Select, or push the join to the source ($expand / FetchXML link-entity).' },
  BP003: { severity: 'warning', title: 'Read call inside a loop', fix: 'Read once outside the loop (one Get items / List rows with $filter), then join with a dictionary lookup, or use $expand on the lookup column. N calls become 1–2.' },
  BP004: { severity: 'warning', title: 'One write per item inside a loop', fix: 'Batch the writes: SharePoint $batch via "Send an HTTP request to SharePoint" (≤ 500 items per batch, ~90 s server limit, parse every part), Dataverse CreateMultiple/UpdateMultiple/UpsertMultiple (100–1,000 per call, via HTTP with Microsoft Entra ID), Excel: one Office Script call with setValues(). Loop over chunk(items, N), not over items.' },
  BP005: { severity: 'error', title: 'Shared variable written in a concurrent loop', fix: 'Parallel iterations overwrite each other (race). Use item()/Compose inside the loop, or Select + an expression after it. If you truly need a variable, set concurrency to 1.' },
  BP006: { severity: 'info', title: 'Sequential loop with independent side effects', fix: 'Apply to each runs one item at a time by default. If order does not matter, set concurrency to 20–50 (stay under the target\'s limits: SharePoint 600 calls/min per connection, Dataverse 52 concurrent requests per user).' },
  BP007: { severity: 'warning', title: 'Reads the whole table without a filter', fix: 'Filter at the source: SharePoint Filter Query ($filter, indexed columns), Dataverse Filter rows / FetchXML. Retrieve only the rows you use.' },
  BP008: { severity: 'info', title: 'Retrieves every column', fix: 'Dataverse: set Select columns ($select). SharePoint: Limit columns by view, or Send an HTTP request with $select (+ $expand for lookups). Payloads drop by 80–99%.' },
  BP009: { severity: 'error', title: 'Pagination off: results silently truncated', fix: 'SharePoint Get items returns only the first 100 items unless pagination is on (Settings > Pagination, threshold sized to your data) or $top is set. Dataverse List rows stops at 5,000.' },
  BP010: { severity: 'warning', title: 'Filter array right after a read', fix: 'Move the condition into the read action\'s $filter / Filter rows so the source returns fewer rows.' },
  BP011: { severity: 'warning', title: 'Flow runs and then terminates instead of using a trigger condition', fix: 'Add a trigger condition (Settings > Trigger conditions), e.g. @equals(triggerOutputs()?[\'body/Status/Value\'], \'Approved\'). Irrelevant events then never start a run and are never billed.' },
  BP012: { severity: 'info', title: 'Re-reads the item that triggered the flow', fix: 'The trigger output already contains the item\'s fields: use triggerOutputs()?[\'body/…\'] instead of a Get item.' },
  BP013: { severity: 'warning', title: 'Response sent after long-running work', fix: 'Power Apps and HTTP callers time out after 120 s and the run is marked Failed. Respond first (status + run id), keep working, and report progress through Dataverse/SharePoint.' },
  BP014: { severity: 'info', title: 'Variable that is never changed', fix: 'Use a Compose: variables are global, locked inside concurrent loops, and each Set is a billed action.' },
  BP015: { severity: 'warning', title: 'SharePoint $batch chunks larger than 500', fix: 'SharePoint processes a $batch serially and stops after ~90 s of server time (measured: 1,000 items → 873–941 created). Use ≤ 500 per batch (250 × 4 in parallel was fastest) and count 2xx/4xx/5xx per part.' },
  BP016: { severity: 'info', title: 'Do until without explicit limits', fix: 'Do until stops silently after 60 iterations or 1 hour by default. Set Limit count / Timeout explicitly to match the data.' },
};

// ------------------------------------------------------------------ input
function findDefinition(obj, depth = 0) {
  if (!obj || depth > 6) return undefined;
  if (typeof obj === 'string') {
    try { return findDefinition(JSON.parse(obj), depth + 1); } catch { return undefined; }
  }
  if (typeof obj !== 'object') return undefined;
  if (obj.actions && typeof obj.actions === 'object' && (obj.triggers || obj.$schema)) return obj;
  for (const k of ['definition', 'properties', 'clientdata', 'flow', 'result', 'data', 'body']) {
    const d = findDefinition(obj[k], depth + 1);
    if (d) return d;
  }
  return undefined;
}

// ------------------------------------------------------------------ helpers
const CONNECTOR_TYPES = new Set(['OpenApiConnection', 'ApiConnection', 'OpenApiConnectionWebhook', 'ApiConnectionWebhook']);
const VAR_WRITE = new Set(['SetVariable', 'AppendToArrayVariable', 'AppendToStringVariable', 'IncrementVariable', 'DecrementVariable']);
const DATA_ONLY = new Set(['Compose', 'If', 'Switch', 'Scope', 'Query', 'Select', 'Join', 'Table', 'ParseJson', ...VAR_WRITE, 'InitializeVariable']);
const LOOPS = new Set(['Foreach', 'Until']);

const str = (v) => (typeof v === 'string' ? v : JSON.stringify(v ?? ''));
const opId = (a) => a?.inputs?.host?.operationId ?? '';
const connector = (a) => {
  const h = a?.inputs?.host ?? {};
  return String(h.apiId ?? h.connectionName ?? h.connection?.name ?? '').split('/').pop().toLowerCase();
};
const params = (a) => a?.inputs?.parameters ?? {};

function httpCall(a) {
  // Returns { method, uri } for HTTP-like actions, else undefined.
  if (a.type === 'Http') return { method: String(a.inputs?.method ?? 'GET').toUpperCase(), uri: str(a.inputs?.uri) };
  if (!CONNECTOR_TYPES.has(a.type)) return undefined;
  const p = params(a);
  if (opId(a) === 'HttpRequest') return { method: String(p['parameters/method'] ?? 'GET').toUpperCase(), uri: str(p['parameters/uri']) };
  if (opId(a) === 'InvokeHttp' || opId(a) === 'InvokeHttpRequest') return { method: String(p['request/method'] ?? p.method ?? 'GET').toUpperCase(), uri: str(p['request/url'] ?? p.url) };
  return undefined;
}

const isBulkUri = (uri) => /\$batch|Multiple\b|Microsoft\.Dynamics\.CRM\.(Create|Update|Upsert|Delete)Multiple/i.test(uri);

// Bulkable single-record writes, per connector.
const WRITE_OPS = {
  sharepointonline: /^(PostItem|PatchItem|DeleteItem|PatchFileItem|PatchFileItemWithPredictedValues|CopyFile|CreateFile)$/,
  commondataserviceforapps: /^(CreateRecord|UpdateRecord|UpdateOnlyRecord|DeleteRecord|UpsertRecord)$/,
  excelonlinebusiness: /^(AddRowV2|AddRow|PatchItem|UpdateRow|DeleteItem|DeleteRow)$/,
  sql: /^(PostItem_V2|PatchItem_V2|DeleteItem_V2|PostItem|PatchItem|DeleteItem)$/,
};
const READ_OP = /^(Get|List|Search|Find|Query)/;

function classify(a) {
  const http = httpCall(a);
  if (http) {
    if (isBulkUri(http.uri)) return 'bulk';
    return http.method === 'GET' ? 'read' : 'write';
  }
  if (!CONNECTOR_TYPES.has(a.type)) return undefined;
  const c = connector(a).replace(/^shared_/, '');
  const op = opId(a);
  if (op === 'ExecuteChangeset') return 'bulk';
  if (WRITE_OPS[c]?.test(op)) return 'write';
  if (READ_OP.test(op)) return 'read';
  return 'other';
}

function children(a) {
  const out = [];
  if (a.actions) out.push(a.actions);
  if (a.else?.actions) out.push(a.else.actions);
  if (a.cases) for (const c of Object.values(a.cases)) if (c?.actions) out.push(c.actions);
  if (a.default?.actions) out.push(a.default.actions);
  return out;
}

function collect(defn) {
  const nodes = new Map();
  const walk = (actions, parent, loops, path) => {
    for (const [name, a] of Object.entries(actions ?? {})) {
      if (!a || typeof a !== 'object') continue;
      const node = { name, a, parent, loops, path: [...path, name], siblings: actions };
      nodes.set(name, node);
      const innerLoops = LOOPS.has(a.type) ? [...loops, name] : loops;
      for (const c of children(a)) walk(c, name, innerLoops, [...path, name]);
    }
  };
  walk(defn.actions, undefined, [], []);
  return nodes;
}

const descendants = (nodes, name) => [...nodes.values()].filter((n) => n.path.includes(name) && n.name !== name);
const refName = (expr) => /^@?\{?\s*(?:outputs|body)\('([^']+)'\)/.exec(String(expr ?? '').trim())?.[1];

// ------------------------------------------------------------------ lint
export function lint(input) {
  const defn = findDefinition(input);
  if (!defn) throw new Error('No flow definition found (expected an object with "triggers" and "actions").');
  const nodes = collect(defn);
  const findings = [];
  const add = (rule, node, detail, severity) => findings.push({ rule, severity: severity ?? RULES[rule].severity, title: RULES[rule].title, action: node?.path.join(' › ') ?? '(trigger)', detail, fix: RULES[rule].fix });

  const triggers = Object.entries(defn.triggers ?? {});
  const allText = JSON.stringify(defn);

  // Loop rules
  for (const n of nodes.values()) {
    const { a } = n;
    if (a.type !== 'Foreach' && a.type !== 'Until') continue;
    const body = descendants(nodes, n.name);
    const reps = Number(a.runtimeConfiguration?.concurrency?.repetitions ?? 1);
    const kinds = body.map((d) => ({ d, k: classify(d.a) }));

    if (a.type === 'Foreach' && body.length && body.every((d) => DATA_ONLY.has(d.a.type))) {
      const agg = body.some((d) => ['IncrementVariable', 'SetVariable'].includes(d.a.type));
      add('BP001', n, `${body.length} data-only action(s) per iteration${agg ? ' (aggregating into a variable: use xpath sum/count or length())' : ''}.`);
    }
    if (n.loops.length) add('BP002', n, `Inside loop "${n.loops.at(-1)}".`);

    for (const { d, k } of kinds) {
      // Do until loops that read page by page are a paging pattern, not a per-item read.
      if (k === 'read' && a.type === 'Foreach' && d.loops.at(-1) === n.name) add('BP003', d, `${opId(d.a) || d.a.type} runs once per iteration of "${n.name}".`);
    }
    const chunked = (() => {
      const src = String(a.foreach ?? '');
      if (/chunk\(/.test(src)) return true;
      const ref = refName(src);
      return ref ? /chunk\(/.test(JSON.stringify(nodes.get(ref)?.a?.inputs ?? '')) : false;
    })();
    if (!chunked) {
      for (const { d, k } of kinds) {
        if (k === 'write' && d.loops.at(-1) === n.name) add('BP004', d, `${opId(d.a) || d.a.type} writes one record per iteration of "${n.name}".`);
      }
    }

    const varWrites = body.filter((d) => VAR_WRITE.has(d.a.type));
    if (a.type === 'Foreach' && reps > 1 && varWrites.length) {
      const sets = varWrites.filter((d) => d.a.type === 'SetVariable');
      add('BP005', n, `Concurrency ${reps} with ${varWrites.map((d) => `${d.a.type} "${d.name}"`).join(', ')}.${sets.length ? '' : ' Append/Increment are serialized by a lock, so the concurrency gain is lost and order is not preserved.'}`);
    }
    if (a.type === 'Foreach' && reps <= 1 && !varWrites.length && kinds.some(({ k }) => k === 'write' || k === 'other' || k === 'read') && !chunked) {
      add('BP006', n, 'Concurrency is 1 (default).');
    }
    if (a.type === 'Until' && !(a.limit?.count || a.limit?.timeout)) add('BP016', n, 'No limit.count / limit.timeout set.');
  }

  // Source rules
  for (const n of nodes.values()) {
    const { a } = n;
    if (!CONNECTOR_TYPES.has(a.type)) continue;
    const c = connector(a).replace(/^shared_/, '');
    const op = opId(a);
    const p = params(a);
    const paged = a.runtimeConfiguration?.paginationPolicy?.minimumItemCount;
    if (c === 'sharepointonline' && op === 'GetItems') {
      if (!p.$filter) add('BP007', n, 'No Filter Query ($filter).');
      if (!p.view && !p.$select) add('BP008', n, 'No view / $select: all columns are returned.');
      if (!paged && !p.$top) add('BP009', n, 'No pagination and no $top: only the first 100 items are returned.');
    }
    if (c === 'commondataserviceforapps' && op === 'ListRecords' && !p.fetchXml) {
      if (!p.$filter) add('BP007', n, 'No Filter rows ($filter).');
      if (!p.$select) add('BP008', n, 'No Select columns ($select).');
      if (!paged && !p.$top) add('BP009', n, 'No pagination and no Row count: at most 5,000 rows are returned.');
    }
  }

  // Filter array right after a read
  for (const n of nodes.values()) {
    if (n.a.type !== 'Query') continue;
    const ref = refName(n.a.inputs?.from);
    const src = ref && nodes.get(ref);
    if (src && classify(src.a) === 'read' && !src.loops.length) {
      const http = httpCall(src.a);
      const filtered = Boolean(params(src.a).$filter || params(src.a).fetchXml || (http && /\$filter=/i.test(http.uri)));
      add('BP010', n, `Filters the output of "${ref}" in the flow${filtered ? ' (the read is already partly filtered: move the remaining condition too, if OData supports it)' : ''}.`, filtered ? 'info' : undefined);
    }
  }

  // Trigger condition vs Terminate
  for (const [tname, t] of triggers) {
    const automated = CONNECTOR_TYPES.has(t.type);
    if (!automated || (Array.isArray(t.conditions) && t.conditions.length)) continue;
    const top = Object.entries(defn.actions ?? {}).map(([name]) => nodes.get(name));
    const ifs = top.filter((x) => x?.a.type === 'If').flatMap((x) => [x, ...descendants(nodes, x.name).filter((d) => d.a.type === 'If')]);
    const terminating = ifs.find((x) => {
      const branches = [x.a.actions ?? {}, x.a.else?.actions ?? {}];
      return branches.some((b) => Object.values(b).some((v) => v?.type === 'Terminate') || Object.keys(b).length === 0);
    });
    if (terminating) add('BP011', terminating, `Trigger "${tname}" has no trigger conditions.`);

    // Get item right after the trigger
    for (const x of top) {
      if (!x || Object.keys(x.a.runAfter ?? {}).length) continue;
      if (classify(x.a) === 'read' && /trigger(Outputs|Body)\(\)/.test(JSON.stringify(params(x.a)))) add('BP012', x, `"${x.name}" reads data the trigger already returned.`);
    }
  }

  // Late response
  for (const n of nodes.values()) {
    if (n.a.type !== 'Response') continue;
    const seen = new Set();
    const ancestors = (name, siblings) => {
      for (const dep of Object.keys(siblings[name]?.runAfter ?? {})) {
        if (seen.has(dep)) continue;
        seen.add(dep);
        ancestors(dep, siblings);
      }
    };
    ancestors(n.name, n.siblings);
    for (let p = n.parent; p; p = nodes.get(p)?.parent) { seen.add(p); ancestors(p, nodes.get(p).siblings); }
    const slow = [...seen].flatMap((name) => [nodes.get(name), ...descendants(nodes, name)]).filter((x) => x && (LOOPS.has(x.a.type) || x.a.type === 'Wait' || x.a.type === 'Workflow'));
    const callerKinds = triggers.map(([, t]) => t.kind);
    const syncCaller = callerKinds.some((k) => ['PowerApp', 'PowerAppV2', 'Http', 'Copilot', 'Skills'].includes(k));
    if (slow.length) add('BP013', n, `${syncCaller ? '' : 'Child flow / other caller: the parent waits for this response. '}Runs after ${[...new Set(slow.map((x) => `${x.a.type} "${x.name}"`))].slice(0, 3).join(', ')}.`, syncCaller ? undefined : 'info');
  }

  // Variables never changed
  for (const n of nodes.values()) {
    if (n.a.type !== 'InitializeVariable') continue;
    for (const v of n.a.inputs?.variables ?? []) {
      const changed = [...nodes.values()].some((x) => VAR_WRITE.has(x.a.type) && x.a.inputs?.name === v.name);
      if (!changed) add('BP014', n, `Variable "${v.name}" is never set, appended or incremented.`);
    }
  }

  // SharePoint $batch chunk size
  if (/_api\/\$batch/.test(allText)) {
    for (const m of allText.matchAll(/chunk\([^()]*(?:\([^()]*\)[^()]*)*,\s*(\d+)\s*\)/g)) {
      if (Number(m[1]) > 500) add('BP015', undefined, `chunk(…, ${m[1]}) used with SharePoint $batch.`);
    }
  }

  // Stats
  const loops = [...nodes.values()].filter((n) => LOOPS.has(n.a.type)).map((n) => ({
    name: n.path.join(' › '),
    type: n.a.type,
    concurrency: Number(n.a.runtimeConfiguration?.concurrency?.repetitions ?? 1),
    actionsPerIteration: descendants(nodes, n.name).length,
  }));
  const order = { error: 0, warning: 1, info: 2 };
  findings.sort((x, y) => order[x.severity] - order[y.severity] || x.rule.localeCompare(y.rule));
  return { stats: { actions: nodes.size, triggers: triggers.map(([n, t]) => `${n} (${t.kind ?? t.type})`), loops }, findings };
}

// ------------------------------------------------------------------ output
export function toMarkdown({ stats, findings }) {
  const icon = { error: '🔴', warning: '🟠', info: '🔵' };
  const lines = [
    `**${stats.actions} actions** · trigger: ${stats.triggers.join(', ') || 'none'} · ${findings.filter((f) => f.severity === 'error').length} error(s), ${findings.filter((f) => f.severity === 'warning').length} warning(s), ${findings.filter((f) => f.severity === 'info').length} info`,
    '',
  ];
  if (stats.loops.length) {
    lines.push('| Loop | Type | Concurrency | Actions per iteration |', '|---|---|---|---|');
    for (const l of stats.loops) lines.push(`| ${l.name} | ${l.type} | ${l.concurrency} | ${l.actionsPerIteration} |`);
    lines.push('', 'Every action inside a loop runs (and is billed) once per item: N items × actions per iteration.', '');
  }
  if (!findings.length) lines.push('✅ No best-practice findings.');
  for (const f of findings) {
    lines.push(`${icon[f.severity]} **${f.rule} ${f.title}** (${f.severity}) — \`${f.action}\``, `   ${f.detail}`, `   Fix: ${f.fix}`, '');
  }
  return lines.join('\n');
}

const isMain = (() => { try { return import.meta.url === pathToFileURL(realpathSync(process.argv[1])).href; } catch { return false; } })();
if (isMain) {
  const args = process.argv.slice(2);
  const file = args.find((x) => !x.startsWith('--'));
  const opt = (k, d) => { const i = args.indexOf(`--${k}`); return i >= 0 ? args[i + 1] : d; };
  if (!file) {
    console.error('Usage: node lint-flow.mjs <flow.json | -> [--format md|json] [--fail-on error|warning|none]');
    process.exit(2);
  }
  let result;
  try {
    const raw = readFileSync(file === '-' ? 0 : file, 'utf8');
    result = lint(JSON.parse(raw));
  } catch (e) {
    console.error(`lint-flow: ${e.message}`);
    process.exit(2);
  }
  console.log(opt('format', 'md') === 'json' ? JSON.stringify(result, null, 2) : toMarkdown(result));
  const failOn = opt('fail-on', 'error');
  const levels = failOn === 'none' ? [] : failOn === 'warning' ? ['error', 'warning'] : ['error'];
  process.exit(result.findings.some((f) => levels.includes(f.severity)) ? 1 : 0);
}
