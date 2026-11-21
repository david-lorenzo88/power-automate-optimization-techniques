// Run: node --test plugins/power-automate-best-practices/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { lint, RULES } from '../skills/flow-best-practices/scripts/lint-flow.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const LAB = join(here, '..', '..', '..', 'lab', 'flows', 'generated');
const rules = (r) => r.findings.map((f) => f.rule).sort();

// ------------------------------------------------------------------ synthetic helpers
const def = (actions, triggers = { manual: { type: 'Request', kind: 'Button', inputs: {} } }) => ({
  $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
  triggers,
  actions,
});
const conn = (connector, operationId, parameters = {}, extra = {}) => ({
  type: 'OpenApiConnection',
  inputs: { host: { apiId: `/providers/Microsoft.PowerApps/apis/${connector}`, connectionName: connector, operationId }, parameters },
  runAfter: {},
  ...extra,
});

test('every rule has a severity, title and fix', () => {
  for (const [id, r] of Object.entries(RULES)) {
    assert.match(id, /^BP\d{3}$/);
    assert.ok(['error', 'warning', 'info'].includes(r.severity), id);
    assert.ok(r.title && r.fix, id);
  }
});

test('accepts raw definitions, clientdata objects and clientdata strings', () => {
  const d = def({ A: { type: 'Compose', inputs: 1, runAfter: {} } });
  assert.equal(lint(d).stats.actions, 1);
  assert.equal(lint({ properties: { definition: d } }).stats.actions, 1);
  assert.equal(lint({ clientdata: JSON.stringify({ properties: { definition: d } }) }).stats.actions, 1);
  assert.throws(() => lint({ hello: 'world' }), /No flow definition/);
});

test('BP007/BP008/BP009: unfiltered, all-column, unpaged reads', () => {
  const r = lint(def({
    Get_items: conn('shared_sharepointonline', 'GetItems', { dataset: 'x', table: 'y' }),
    List_rows: conn('shared_commondataserviceforapps', 'ListRecords', { entityName: 'accounts' }),
  }));
  assert.deepEqual(rules(r), ['BP007', 'BP007', 'BP008', 'BP008', 'BP009', 'BP009']);
  const ok = lint(def({
    Get_items: conn('shared_sharepointonline', 'GetItems', { dataset: 'x', table: 'y', $filter: "Status eq 'Open'", view: 'v' }, { runtimeConfiguration: { paginationPolicy: { minimumItemCount: 5000 } } }),
    List_rows: conn('shared_commondataserviceforapps', 'ListRecords', { entityName: 'accounts', $filter: 'statecode eq 0', $select: 'name', $top: 50 }),
  }));
  assert.deepEqual(rules(ok), []);
});

test('BP004 is not raised when the loop iterates chunks', () => {
  const write = conn('shared_sharepointonline', 'PostItem', { dataset: 'x', table: 'y' });
  const perItem = lint(def({ Loop: { type: 'Foreach', foreach: "@body('Rows')", actions: { Create: write }, runAfter: {} } }));
  assert.ok(rules(perItem).includes('BP004'));
  const chunked = lint(def({
    Chunks: { type: 'Compose', inputs: "@chunk(body('Rows'), 500)", runAfter: {} },
    Loop: { type: 'Foreach', foreach: "@outputs('Chunks')", actions: { Create: write }, runAfter: { Chunks: ['Succeeded'] } },
  }));
  assert.ok(!rules(chunked).includes('BP004'));
});

test('BP011 terminate pattern and BP012 get-after-trigger on automated triggers', () => {
  const trigger = { When_modified: conn('shared_sharepointonline', 'GetOnUpdatedItems', { dataset: 'x', table: 'y' }) };
  const r = lint(def({
    Get_item: conn('shared_sharepointonline', 'GetItem', { dataset: 'x', table: 'y', id: "@triggerOutputs()?['body/ID']" }),
    Check: { type: 'If', expression: {}, actions: {}, else: { actions: { Stop: { type: 'Terminate', inputs: { runStatus: 'Cancelled' } } } }, runAfter: { Get_item: ['Succeeded'] } },
  }, trigger));
  assert.deepEqual(rules(r), ['BP011', 'BP012']);
  trigger.When_modified.conditions = [{ expression: "@equals(triggerOutputs()?['body/Status/Value'], 'Approved')" }];
  assert.ok(!rules(lint(def({ Check: { type: 'If', expression: {}, actions: {}, runAfter: {} } }, trigger))).includes('BP011'));
});

test('BP013 late response, BP014 constant variable, BP015 batch size, BP016 until limits', () => {
  const r = lint(def({
    Init: { type: 'InitializeVariable', inputs: { variables: [{ name: 'Url', type: 'string', value: 'x' }] }, runAfter: {} },
    Loop: { type: 'Until', expression: '@true', actions: { W: { type: 'Wait', inputs: { interval: { count: 1, unit: 'Second' } } } }, runAfter: { Init: ['Succeeded'] } },
    Parts: { type: 'Compose', inputs: "@chunk(body('Rows'), 1000)", runAfter: { Loop: ['Succeeded'] } },
    Send: conn('shared_sharepointonline', 'HttpRequest', { 'parameters/method': 'POST', 'parameters/uri': '_api/$batch' }, { runAfter: { Parts: ['Succeeded'] } }),
    Respond: { type: 'Response', kind: 'PowerApp', inputs: { statusCode: 200 }, runAfter: { Send: ['Succeeded'] } },
  }, { manual: { type: 'Request', kind: 'PowerAppV2', inputs: {} } }));
  assert.deepEqual(rules(r), ['BP013', 'BP014', 'BP015', 'BP016']);
});

// ------------------------------------------------------------------ measured lab flows (repo checkout only)
const labExpect = {
  'UC1-SLOW': ['BP001'], 'UC2-SLOW': ['BP003'], 'UC3-SLOW': ['BP001', 'BP002'], 'UC4-SLOW': ['BP004'],
  'UC5-SLOW': ['BP004'], 'UC6-SLOW': ['BP003'], 'UC7-SLOW': ['BP011'], 'UC8-SLOW': ['BP001', 'BP005', 'BP009'], 'UC9-SLOW': ['BP004'],
};
const fastFlows = ['UC1-FAST', 'UC2-FAST', 'UC3-FAST', 'UC4-FAST', 'UC5-FAST', 'UC6-FAST', 'UC6DV-FAST', 'UC7-FAST', 'UC8-FAST', 'UC9-FAST', 'UC10-FAST'];
const load = (n) => JSON.parse(readFileSync(join(LAB, `${n}.json`), 'utf8'));

test('lab SLOW flows are flagged with the anti-pattern they demonstrate', { skip: !existsSync(LAB) && 'lab fixtures not present' }, () => {
  for (const [name, expected] of Object.entries(labExpect)) {
    const got = rules(lint(load(name)));
    for (const id of expected) assert.ok(got.includes(id), `${name}: expected ${id}, got ${got.join(',')}`);
  }
});

test('lab FAST flows have no errors or warnings', { skip: !existsSync(LAB) && 'lab fixtures not present' }, () => {
  for (const name of fastFlows) {
    const bad = lint(load(name)).findings.filter((f) => f.severity !== 'info');
    assert.deepEqual(bad.map((f) => `${f.rule} ${f.action}`), [], name);
  }
});
