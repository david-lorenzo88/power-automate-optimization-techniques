// Runs lab flows from the command line (through the LAB-TestRunner flow) and waits for
// their pa_benchmarkrun rows. Same telemetry as a run from the Code App.
//
//   node lab/scripts/30-run.mjs UC1-FAST --volume 500
//   node lab/scripts/30-run.mjs --pair UC1 --volume 500          all variants of UC1, in parallel, compared
//   node lab/scripts/30-run.mjs SEED-SP --timeout 3600
//   node lab/scripts/30-run.mjs RESET --payload '{"target":"uc4"}'
//   options: --label <text>  --payload '<json merged into the payload>'  --timeout <s> (default 900)  --no-wait
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { dvOk, PREFIX } from '../../poc/scripts/lib/dv.mjs';
import { azToken } from './lib/sp.mjs';

const P = PREFIX;
const here = dirname(fileURLToPath(import.meta.url));
const manifest = JSON.parse(readFileSync(join(here, '..', 'flows', 'deployed.json'), 'utf8'));
const cfg = JSON.parse(readFileSync(join(here, '..', 'lab.config.json'), 'utf8'));

const argv = process.argv.slice(2);
const opt = (name, def) => { const i = argv.indexOf(`--${name}`); return i >= 0 ? argv[i + 1] : def; };
const flag = (name) => argv.includes(`--${name}`);
const valued = new Set(['volume', 'label', 'payload', 'timeout', 'pair']);
const positional = argv.filter((a, i) => !a.startsWith('--') && !(i > 0 && valued.has(argv[i - 1].slice(2))));

let names = positional;
if (opt('pair')) {
  const uc = opt('pair').toUpperCase();
  names = Object.entries(manifest).filter(([, v]) => v.app && v.scenario === uc).map(([k]) => k)
    .sort((a, b) => ['SLOW', 'MID', 'FAST'].indexOf(manifest[a].variant) - ['SLOW', 'MID', 'FAST'].indexOf(manifest[b].variant));
}
if (!names.length) { console.error('Usage: 30-run.mjs <FLOW...> | --pair <UCx> [--volume N]'); process.exit(1); }
for (const n of names) if (!manifest[n]) { console.error(`Unknown flow ${n}. Known: ${Object.keys(manifest).join(', ')}`); process.exit(1); }

const runner = manifest['LAB-TestRunner'];
if (!runner?.flowId) { console.error('LAB-TestRunner is not deployed (run 20-deploy-flows.mjs).'); process.exit(1); }
const flowApi = `https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/environments/${cfg.environmentId}/flows/${runner.flowId}`;
const cb = await fetch(`${flowApi}/triggers/manual/listCallbackUrl?api-version=2016-11-01`, { method: 'POST', headers: { Authorization: `Bearer ${azToken('https://service.flow.microsoft.com/')}` } });
if (!cb.ok) { console.error(`listCallbackUrl -> ${cb.status}: ${await cb.text()}`); process.exit(1); }
const callbackUrl = (await cb.json()).response.value; // contains a SAS signature: never print or store it

const extra = opt('payload') ? JSON.parse(opt('payload')) : {};
const label = opt('label', 'cli');
const runs = [];
for (const name of names) {
  const correlationId = randomUUID();
  const payload = { correlationId, label, ...(opt('volume') ? { volume: Number(opt('volume')) } : {}), ...extra };
  const r = await fetch(callbackUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ flow: name, payload }) });
  if (!r.ok) { console.error(`${name}: runner HTTP ${r.status} ${await r.text()}`); continue; }
  console.log(`▶️  ${name.padEnd(12)} correlationId ${correlationId}`);
  runs.push({ name, correlationId, started: Date.now() });
}
if (flag('no-wait')) process.exit(0);

const cols = ['pa_status', 'pa_volume', 'pa_durationms', 'pa_recordsprocessed', 'pa_recordsretrieved', 'pa_loopiterations', 'pa_connectorcalls', 'pa_actionsestimated', 'pa_payloadbytes', 'pa_resultchecksum', 'pa_runurl', 'pa_error', 'pa_details'];
const timeout = Number(opt('timeout', 900)) * 1000;
const rows = {};
while (runs.some((r) => !rows[r.name] || rows[r.name].pa_status === 'Running')) {
  for (const r of runs) {
    if (rows[r.name] && rows[r.name].pa_status !== 'Running') continue;
    const res = await dvOk('GET', `${P}_benchmarkruns?$select=${cols.join(',')}&$filter=${P}_name eq '${r.correlationId}'`);
    const row = res.json.value[0];
    if (row) rows[r.name] = row;
    if (Date.now() - r.started > timeout) rows[r.name] = { ...(row || {}), pa_status: row ? `${row.pa_status} (timed out waiting)` : 'not started (timed out waiting)' };
  }
  if (runs.some((r) => !rows[r.name] || rows[r.name].pa_status === 'Running')) {
    process.stdout.write(`\r⏳ ${runs.map((r) => `${r.name}: ${rows[r.name]?.pa_status ?? 'queued'}`).join(' · ')} (${Math.round((Date.now() - runs[0].started) / 1000)} s)   `);
    await new Promise((res) => setTimeout(res, 5000));
  }
}
process.stdout.write('\n');

const table = runs.map((r) => {
  const x = rows[r.name] || {};
  return {
    flow: r.name, status: x.pa_status, volume: x.pa_volume, ms: x.pa_durationms, processed: x.pa_recordsprocessed, retrieved: x.pa_recordsretrieved,
    loops: x.pa_loopiterations, calls: x.pa_connectorcalls, actions: x.pa_actionsestimated, payload: x.pa_payloadbytes, checksum: x.pa_resultchecksum,
  };
});
console.table(table);
for (const r of runs) {
  const x = rows[r.name] || {};
  if (x.pa_error) console.log(`❌ ${r.name} error: ${x.pa_error.slice(0, 1500)}`);
  if (x.pa_runurl) console.log(`🔗 ${r.name}: ${x.pa_runurl}`);
}
if (runs.length > 1) {
  const fast = table.find((t) => t.flow.endsWith('FAST'));
  for (const t of table.filter((t) => t !== fast)) {
    if (fast?.ms && t.ms) console.log(`⚡ ${fast.flow} is ${(t.ms / fast.ms).toFixed(1)}× faster than ${t.flow}; checksum ${t.checksum === fast.checksum ? '✅ match' : `❌ differs (${t.checksum} vs ${fast.checksum})`}`);
  }
}
