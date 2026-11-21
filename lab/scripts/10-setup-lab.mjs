// Creates the demo data stores for the slow-vs-fast use cases (docs/02-demo-scenarios.md)
// and writes lab/lab.config.json for the flow generator. Idempotent: re-run any time.
//
//   node lab/scripts/10-setup-lab.mjs
//
// Needs: DATAVERSE_URL, `az login`, the PoC setup (poc/scripts/01-setup-dataverse.mjs)
// and a bound + healthy `FlowOptLab SharePoint` connection reference.
// Optional: SP_SITE_URL (default: tenant root site).
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { dv, dvOk, label, solutionHeader, ORG, PREFIX, SOLUTION } from '../../poc/scripts/lib/dv.mjs';
import { ensureFields, ensureList, sp, spOk } from './lib/sp.mjs';

const P = PREFIX;
const here = dirname(fileURLToPath(import.meta.url));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ------------------------------------------------------------------ Dataverse
const req = { Value: 'None', CanBeChanged: true, ManagedPropertyLogicalName: 'canmodifyrequirementlevelsettings' };
const base = (schema, display) => ({ SchemaName: schema, DisplayName: label(display), Description: label(display), RequiredLevel: req });
const text = (schema, display, max = 100) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata', ...base(schema, display), MaxLength: max, FormatName: { Value: 'Text' } });
const dec = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.DecimalAttributeMetadata', ...base(schema, display), Precision: 2, MinValue: -100000000000, MaxValue: 100000000000 });
const dt = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata', ...base(schema, display), Format: 'DateOnly', DateTimeBehavior: { Value: 'DateOnly' } });

// Metadata writes right after a create can 404 (not readable yet) or 400 ("unexpected error"); retry those.
async function metaWrite(method, path, body) {
  for (let i = 0; ; i++) {
    const r = await dv(method, path, body, solutionHeader());
    if (r.ok) return r;
    if (i < 6 && (r.status === 404 || (r.status === 400 && /unexpected error/i.test(r.text)))) { await sleep(5000); continue; }
    throw new Error(`${method} ${path} -> ${r.status}: ${r.json?.error?.message || r.text}`);
  }
}

async function ensureTable(schema, display, plural, primary, columns) {
  const logical = schema.toLowerCase();
  const exists = await dv('GET', `EntityDefinitions(LogicalName='${logical}')?$select=LogicalName`);
  if (!exists.ok) {
    await metaWrite('POST', 'EntityDefinitions', {
      '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata', SchemaName: schema,
      DisplayName: label(display), DisplayCollectionName: label(plural), Description: label(`${display} (Flow Optimization Lab)`),
      OwnershipType: 'UserOwned', IsActivity: false, HasActivities: false, HasNotes: false,
      Attributes: [{ ...primary, IsPrimaryName: true }],
    });
    console.log(`✅ table ${logical} created`);
  } else console.log(`ℹ️  table ${logical} exists`);
  let attrs;
  for (let i = 0; i < 30; i++) {
    attrs = await dv('GET', `EntityDefinitions(LogicalName='${logical}')/Attributes?$select=LogicalName`);
    if (attrs.ok) break;
    await sleep(2000);
  }
  const have = new Set(attrs.json.value.map((a) => a.LogicalName));
  for (const col of columns) {
    if (have.has(col.SchemaName.toLowerCase())) continue;
    await metaWrite('POST', `EntityDefinitions(LogicalName='${logical}')/Attributes`, col);
    console.log(`   + ${col.SchemaName}`);
  }
}

async function ensureConnRef(logicalName, displayName, connector) {
  const r = await dvOk('GET', `connectionreferences?$select=connectionid&$filter=connectionreferencelogicalname eq '${logicalName}'`);
  if (r.json.value.length) {
    return console.log(`ℹ️  connection reference ${logicalName}: ${r.json.value[0].connectionid ? 'bound' : 'NOT bound yet'}`);
  }
  await dvOk('POST', 'connectionreferences', {
    connectionreferencelogicalname: logicalName, connectionreferencedisplayname: displayName,
    connectorid: `/providers/Microsoft.PowerApps/apis/${connector}`,
  }, solutionHeader());
  console.log(`✅ connection reference ${logicalName} created → bind it in the solution UI`);
}

await ensureTable(`${P}_order`, 'Order', 'Orders', text(`${P}_name`, 'Order ID', 20), [
  text(`${P}_customerid`, 'Customer ID', 20),
  text(`${P}_region`, 'Region', 20),
  text(`${P}_status`, 'Status', 20),
  dec(`${P}_amount`, 'Amount'),
  dt(`${P}_orderdate`, 'Order date'),
]);
await ensureConnRef(`${P}_httpdataverse`, 'FlowOptLab HTTP Dataverse (Entra ID)', 'shared_webcontents');
await ensureConnRef(`${P}_excel`, 'FlowOptLab Excel', 'shared_excelonlinebusiness');
const pub = await dv('POST', 'PublishAllXml', {});
console.log(pub.ok ? '✅ published customizations' : `⚠️ publish: ${pub.status}`);

// ------------------------------------------------------------------ SharePoint
const ctx = await sp();
console.log(`\nSharePoint site: ${ctx.site}`);

const customers = await ensureList('Customers', 'Flow Optimization Lab: lookup target (2,000 items)');
console.log(`${customers.created ? '✅' : 'ℹ️ '} list Customers (${customers.ItemCount} items)`);
console.log('   +', (await ensureFields('Customers', [
  { name: 'CustomerName', kind: 'text' },
  { name: 'Segment', kind: 'text' },
  { name: 'Country', kind: 'text' },
])).join(', ') || 'no new fields');

const orders = await ensureList('Orders', 'Flow Optimization Lab: 20,000 orders with ~20 noise columns');
console.log(`${orders.created ? '✅' : 'ℹ️ '} list Orders (${orders.ItemCount} items)`);
const noise = Array.from({ length: 20 }, (_, i) => ({ name: `Noise${String(i + 1).padStart(2, '0')}`, kind: 'text' }));
console.log('   +', (await ensureFields('Orders', [
  // Not 'CustomerId': SharePoint REST exposes the `Customer` lookup's id as CustomerId.
  { name: 'CustCode', kind: 'text', indexed: true },
  { name: 'Customer', kind: 'lookup', lookupListId: customers.Id },
  { name: 'Region', kind: 'text', indexed: true },
  { name: 'Status', kind: 'text', indexed: true },
  { name: 'Amount', kind: 'number' },
  { name: 'OrderDate', kind: 'date', indexed: true },
  ...noise,
])).join(', ') || 'no new fields');

const orderImport = await ensureList('OrderImport', 'Flow Optimization Lab: bulk create target (UC4)');
console.log(`${orderImport.created ? '✅' : 'ℹ️ '} list OrderImport (${orderImport.ItemCount} items)`);
console.log('   +', (await ensureFields('OrderImport', [
  { name: 'Amount', kind: 'number' },
  { name: 'Region', kind: 'text' },
  { name: 'BatchTag', kind: 'text', indexed: true },
])).join(', ') || 'no new fields');

const uc7 = await ensureList('UC7Items', 'Flow Optimization Lab: trigger condition demo (UC7)');
console.log(`${uc7.created ? '✅' : 'ℹ️ '} list UC7Items (${uc7.ItemCount} items)`);
console.log('   +', (await ensureFields('UC7Items', [
  { name: 'Status', kind: 'text' },
  { name: 'Notes', kind: 'text' },
  { name: 'CorrelationId', kind: 'text' },
])).join(', ') || 'no new fields');

// Field internal names can differ from titles (e.g. a lookup); record what SharePoint actually uses.
const fieldNames = async (list) => Object.fromEntries((await spOk('GET', `_api/web/lists/getbytitle('${list}')/fields?$select=Title,InternalName&$filter=Hidden eq false&$top=500`)).json.value.map((f) => [f.Title, f.InternalName]));

const reread = async (t) => (await spOk('GET', `_api/web/lists/getbytitle('${t}')?$select=Id,ListItemEntityTypeFullName`)).json;
const lists = {};
for (const t of ['Customers', 'Orders', 'OrderImport', 'UC7Items']) {
  const l = await reread(t);
  lists[t] = { id: l.Id, entityType: l.ListItemEntityTypeFullName };
}
const ordersFields = await fieldNames('Orders');

const config = {
  generatedAt: new Date().toISOString(),
  environmentId: ctx.env,
  dataverseUrl: ORG,
  solution: SOLUTION,
  prefix: P,
  siteUrl: ctx.site,
  lists,
  customerLookupField: ordersFields.Customer,
};
const file = join(here, '..', 'lab.config.json');
writeFileSync(file, JSON.stringify(config, null, 2) + '\n');
console.log(`\n📝 wrote ${file}`);
console.log(JSON.stringify(config, null, 2));
