// Creates/updates the lab flows (lab/flows/scenarios.mjs) inside the solution via the
// Dataverse Web API and turns them on. Two phases: flows that call child flows
// (UC10-FAST, LAB-TestRunner) need the child's workflowid, so they are built after the rest.
//
//   node lab/scripts/20-deploy-flows.mjs                 deploy everything
//   node lab/scripts/20-deploy-flows.mjs UC1-SLOW UC1-FAST   only these (plus nothing else)
//   node lab/scripts/20-deploy-flows.mjs --dry-run       only write lab/flows/generated/*.json
//
// Writes lab/flows/deployed.json: { name: { workflowid, flowId, state, app } }.
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { REF_FOR_API } from '../flows/lib.mjs';
import { buildLabFlows, clientData } from '../flows/scenarios.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
const cfg = JSON.parse(readFileSync(join(root, 'lab.config.json'), 'utf8'));
const args = process.argv.slice(2);
const dryRun = args.includes('--dry-run');
const only = args.filter((a) => !a.startsWith('--'));
const refs = REF_FOR_API(cfg.prefix);
const genDir = join(root, 'flows', 'generated');
mkdirSync(genDir, { recursive: true });
const manifestFile = join(root, 'flows', 'deployed.json');
const manifest = existsSync(manifestFile) ? JSON.parse(readFileSync(manifestFile, 'utf8')) : {};

const write = (f) => writeFileSync(join(genDir, `${f.name}.json`), JSON.stringify(clientData(f, refs), null, 2) + '\n');

if (dryRun) {
  let fake = {};
  let all = buildLabFlows(cfg, fake);
  while (all.some((f) => !fake[f.name])) {
    fake = Object.fromEntries(all.map((f) => [f.name, '00000000-0000-0000-0000-000000000000']));
    all = buildLabFlows(cfg, fake);
  }
  all.forEach(write);
  console.log(`📝 wrote ${all.length} definitions to lab/flows/generated/`);
  process.exit(0);
}

const { dv, dvOk, solutionHeader } = await import('../../poc/scripts/lib/dv.mjs');

async function deploy(f) {
  write(f);
  const body = { clientdata: JSON.stringify(clientData(f, refs)), description: f.description };
  const existing = await dvOk('GET', `workflows?$select=workflowid,statecode&$filter=name eq '${f.name}' and category eq 5`);
  let id;
  if (existing.json.value.length) {
    id = existing.json.value[0].workflowid;
    if (existing.json.value[0].statecode === 1) await dv('PATCH', `workflows(${id})`, { statecode: 0, statuscode: 1 });
    await dvOk('PATCH', `workflows(${id})`, body);
  } else {
    id = (await dvOk('POST', 'workflows', { category: 5, type: 1, primaryentity: 'none', name: f.name, ...body },
      { ...solutionHeader(), Prefer: 'return=representation' })).json.workflowid;
  }
  const on = await dv('PATCH', `workflows(${id})`, { statecode: 1, statuscode: 2 });
  const info = (await dvOk('GET', `workflows(${id})?$select=resourceid,statecode`)).json;
  manifest[f.name] = { workflowid: id, flowId: info.resourceid || manifest[f.name]?.flowId || null, state: info.statecode === 1 ? 'on' : 'off', app: !!f.meta.app, scenario: f.meta.scenario, variant: f.meta.variant };
  const msg = on.ok ? '▶️  on' : `⏸️  off: ${(on.json?.error?.message || on.text).slice(0, 300)}`;
  console.log(`${existing.json.value.length ? '🔁' : '✅'} ${f.name.padEnd(16)} ${msg}`);
  return id;
}

const ids = Object.fromEntries(Object.entries(manifest).map(([k, v]) => [k, v.workflowid]));
// Phases: each build can add flows whose child flows now have ids (UC10-FAST, then LAB-TestRunner).
const done = new Set();
for (let phase = 1; phase <= 4; phase++) {
  const runnable = Object.entries(manifest).filter(([, v]) => v.state === 'on').map(([k]) => k);
  const todo = buildLabFlows(cfg, ids, { runnable }).filter((f) => !done.has(f.name));
  if (!todo.length) break;
  for (const f of todo) {
    done.add(f.name);
    if (only.length && !only.includes(f.name)) continue;
    ids[f.name] = await deploy(f);
  }
}
writeFileSync(manifestFile, JSON.stringify(manifest, null, 2) + '\n');
const off = Object.entries(manifest).filter(([, v]) => v.state !== 'on').map(([k]) => k);
console.log(`\n📝 lab/flows/deployed.json (${Object.keys(manifest).length} flows)${off.length ? `\n⚠️  not on: ${off.join(', ')}` : ''}`);
