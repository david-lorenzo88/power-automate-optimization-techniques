// Checklist #2 (flows): create/update the PoC cloud flows inside the solution
// through the Dataverse Web API (workflow table, category 5 = modern flow),
// then try to turn them on.
//
//   node poc/scripts/02-create-poc-flows.mjs            create/update + activate
//   node poc/scripts/02-create-poc-flows.mjs --dry-run  only write JSON to poc/flows/generated/
//
// Turning on fails if the connection references are not bound to connections yet:
// bind them in the solution UI (Connection references) and re-run.
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const dryRun = process.argv.includes('--dry-run');
const here = dirname(fileURLToPath(import.meta.url));
const { buildFlows } = await import('../flows/definitions.mjs');
const prefix = process.env.PUBLISHER_PREFIX || 'pa';
const flows = buildFlows(prefix);

const outDir = join(here, '..', 'flows', 'generated');
mkdirSync(outDir, { recursive: true });
for (const f of flows) writeFileSync(join(outDir, `${f.name}.json`), JSON.stringify(f.clientdata, null, 2));
console.log(`📝 wrote ${flows.length} definitions to poc/flows/generated/`);
if (dryRun) process.exit(0);

const { dv, dvOk, solutionHeader, ORG } = await import('./lib/dv.mjs');

const env = await dv('GET', 'RetrieveCurrentOrganization(AccessType=@p1)?@p1=Microsoft.Dynamics.CRM.EndpointAccessType\'Default\'');
const environmentId = env.json?.Detail?.EnvironmentId;

for (const f of flows) {
  const existing = await dvOk('GET', `workflows?$select=workflowid,statecode,resourceid&$filter=name eq '${f.name}' and category eq 5`);
  let id;
  if (existing.json.value.length) {
    id = existing.json.value[0].workflowid;
    if (existing.json.value[0].statecode === 1) await dv('PATCH', `workflows(${id})`, { statecode: 0, statuscode: 1 });
    await dvOk('PATCH', `workflows(${id})`, { clientdata: JSON.stringify(f.clientdata), description: f.description });
    console.log(`🔁 updated ${f.name}`);
  } else {
    const c = await dvOk('POST', 'workflows', {
      category: 5,
      type: 1,
      primaryentity: 'none',
      name: f.name,
      description: f.description,
      clientdata: JSON.stringify(f.clientdata),
    }, { ...solutionHeader(), Prefer: 'return=representation' });
    id = c.json.workflowid;
    console.log(`✅ created ${f.name}`);
  }

  const on = await dv('PATCH', `workflows(${id})`, { statecode: 1, statuscode: 2 });
  const info = await dvOk('GET', `workflows(${id})?$select=resourceid,statecode`);
  const flowId = info.json.resourceid || '(assigned after first activation)';
  console.log(on.ok ? '   ▶️  turned on' : `   ⏸️  could not turn on (${on.status}): ${on.json?.error?.message ?? on.text}`);
  console.log(`   workflowid (Dataverse): ${id}`);
  console.log(`   flow id (resourceid, used by pa CLI + URLs): ${flowId}`);
  if (environmentId && info.json.resourceid) {
    console.log(`   https://make.powerautomate.com/environments/${environmentId}/solutions/~preferred/flows/${info.json.resourceid}`);
  }
}

console.log(`\nOrg: ${ORG}${environmentId ? `  Environment: ${environmentId}` : ''}`);
console.log('Next: in /app run `pa app list-flows --search Poc` and `pa app add flow --flow-id <flow id>` for each flow.');
