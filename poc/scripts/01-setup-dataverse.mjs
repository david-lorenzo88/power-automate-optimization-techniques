// Checklist #2: publisher + solution + telemetry table (pa_benchmarkrun)
// + bulk test table (pa_orderimport) + connection references.
// Idempotent: safe to run again, it only creates what is missing.
import { dv, dvOk, label, solutionHeader, PREFIX, SOLUTION } from './lib/dv.mjs';

const P = PREFIX;
const PUBLISHER = 'flowoptlab';

// ---------------------------------------------------------------- publisher + solution
async function ensurePublisher() {
  const r = await dvOk('GET', `publishers?$select=publisherid&$filter=uniquename eq '${PUBLISHER}'`);
  if (r.json.value.length) return r.json.value[0].publisherid;
  const c = await dvOk('POST', 'publishers', {
    uniquename: PUBLISHER,
    friendlyname: 'Flow Optimization Lab',
    customizationprefix: P,
    customizationoptionvalueprefix: 72_700,
  }, { Prefer: 'return=representation' });
  console.log('✅ publisher created');
  return c.json.publisherid;
}

async function ensureSolution(publisherId) {
  const r = await dvOk('GET', `solutions?$select=solutionid&$filter=uniquename eq '${SOLUTION}'`);
  if (r.json.value.length) return console.log(`ℹ️  solution ${SOLUTION} exists`);
  await dvOk('POST', 'solutions', {
    uniquename: SOLUTION,
    friendlyname: 'Flow Optimization Lab',
    version: '1.0.0.0',
    'publisherid@odata.bind': `/publishers(${publisherId})`,
  });
  console.log(`✅ solution ${SOLUTION} created`);
}

// ---------------------------------------------------------------- column builders
const req = { Value: 'None', CanBeChanged: true, ManagedPropertyLogicalName: 'canmodifyrequirementlevelsettings' };
const base = (schema, display) => ({ SchemaName: schema, DisplayName: label(display), Description: label(display), RequiredLevel: req });
const text = (schema, display, max = 200) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata', ...base(schema, display), MaxLength: max, FormatName: { Value: 'Text' } });
const memo = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.MemoAttributeMetadata', ...base(schema, display), MaxLength: 100000, Format: 'TextArea' });
const int = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.IntegerAttributeMetadata', ...base(schema, display), Format: 'None', MinValue: -2147483648, MaxValue: 2147483647 });
const dec = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.DecimalAttributeMetadata', ...base(schema, display), Precision: 2, MinValue: -100000000000, MaxValue: 100000000000 });
const dt = (schema, display) => ({ '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata', ...base(schema, display), Format: 'DateAndTime', DateTimeBehavior: { Value: 'UserLocal' } });

async function ensureTable(schema, display, plural, primary, columns) {
  const logical = schema.toLowerCase();
  const exists = await dv('GET', `EntityDefinitions(LogicalName='${logical}')?$select=LogicalName,EntitySetName`);
  if (!exists.ok) {
    await dvOk('POST', 'EntityDefinitions', {
      '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata',
      SchemaName: schema,
      DisplayName: label(display),
      DisplayCollectionName: label(plural),
      Description: label(`${display} (Flow Optimization Lab)`),
      OwnershipType: 'UserOwned',
      IsActivity: false,
      HasActivities: false,
      HasNotes: false,
      Attributes: [{ ...primary, IsPrimaryName: true }],
    }, solutionHeader());
    console.log(`✅ table ${logical} created`);
    // New entity metadata can take a few seconds to become readable (404 right after create).
    // Poll the Attributes endpoint itself: the entity GET can succeed before it does.
    for (let i = 0; i < 30; i++) {
      if ((await dv('GET', `EntityDefinitions(LogicalName='${logical}')/Attributes?$select=LogicalName`)).ok) break;
      await new Promise((res) => setTimeout(res, 2000));
    }
  } else {
    console.log(`ℹ️  table ${logical} exists (entity set: ${exists.json.EntitySetName})`);
  }
  const attrs = await dvOk('GET', `EntityDefinitions(LogicalName='${logical}')/Attributes?$select=LogicalName`);
  const have = new Set(attrs.json.value.map((a) => a.LogicalName));
  for (const col of columns) {
    if (have.has(col.SchemaName.toLowerCase())) continue;
    await dvOk('POST', `EntityDefinitions(LogicalName='${logical}')/Attributes`, col, solutionHeader());
    console.log(`   + ${col.SchemaName}`);
  }
}

// ---------------------------------------------------------------- tables
const benchmarkColumns = [
  text(`${P}_scenario`, 'Scenario', 50),
  text(`${P}_variant`, 'Variant', 50),
  text(`${P}_technique`, 'Technique', 200),
  text(`${P}_status`, 'Status', 50),
  text(`${P}_label`, 'Label', 200),
  int(`${P}_volume`, 'Volume'),
  dt(`${P}_starttime`, 'Start time'),
  dt(`${P}_endtime`, 'End time'),
  int(`${P}_durationms`, 'Duration (ms)'),
  int(`${P}_recordsprocessed`, 'Records processed'),
  int(`${P}_recordsretrieved`, 'Records retrieved'),
  int(`${P}_loopiterations`, 'Loop iterations'),
  int(`${P}_connectorcalls`, 'Connector calls'),
  int(`${P}_actionsestimated`, 'Actions (estimated)'),
  int(`${P}_payloadbytes`, 'Payload bytes'),
  text(`${P}_resultchecksum`, 'Result checksum', 500),
  text(`${P}_runid`, 'Run ID', 100),
  text(`${P}_flowid`, 'Flow ID', 100),
  text(`${P}_environmentid`, 'Environment ID', 100),
  text(`${P}_runurl`, 'Run URL', 1000),
  text(`${P}_runurlsolution`, 'Run URL (solution path)', 1000),
  text(`${P}_editurl`, 'Edit URL', 1000),
  memo(`${P}_details`, 'Details (JSON)'),
  memo(`${P}_error`, 'Error'),
];

const orderImportColumns = [
  text(`${P}_externalid`, 'External ID', 100),
  text(`${P}_region`, 'Region', 50),
  dec(`${P}_amount`, 'Amount'),
  text(`${P}_batchtag`, 'Batch tag', 100),
];

async function ensureAltKey() {
  const logical = `${P}_orderimport`;
  const keys = await dvOk('GET', `EntityDefinitions(LogicalName='${logical}')/Keys?$select=LogicalName`);
  if (keys.json.value.some((k) => k.LogicalName === `${P}_externalid_key`)) return console.log('ℹ️  alternate key exists');
  await dvOk('POST', `EntityDefinitions(LogicalName='${logical}')/Keys`, {
    SchemaName: `${P}_externalid_key`,
    DisplayName: label('External ID key'),
    KeyAttributes: [`${P}_externalid`],
  }, solutionHeader());
  console.log('✅ alternate key on pa_externalid requested (index builds asynchronously, wait ~1-5 min before UpsertMultiple)');
}

// ---------------------------------------------------------------- connection references
async function ensureConnRef(logicalName, displayName, connector) {
  const r = await dvOk('GET', `connectionreferences?$select=connectionreferenceid,connectionid&$filter=connectionreferencelogicalname eq '${logicalName}'`);
  if (r.json.value.length) {
    const bound = r.json.value[0].connectionid ? 'bound to a connection' : 'NOT bound to a connection yet';
    return console.log(`ℹ️  connection reference ${logicalName} exists (${bound})`);
  }
  await dvOk('POST', 'connectionreferences', {
    connectionreferencelogicalname: logicalName,
    connectionreferencedisplayname: displayName,
    connectorid: `/providers/Microsoft.PowerApps/apis/${connector}`,
  }, solutionHeader());
  console.log(`✅ connection reference ${logicalName} created (bind it to a connection in the solution UI)`);
}

// ---------------------------------------------------------------- run
const publisherId = await ensurePublisher();
await ensureSolution(publisherId);

await ensureTable(`${P}_benchmarkrun`, 'Benchmark run', 'Benchmark runs',
  text(`${P}_name`, 'Correlation ID', 100), benchmarkColumns);

await ensureTable(`${P}_orderimport`, 'Order import', 'Order imports',
  text(`${P}_name`, 'Name', 200), orderImportColumns);
await ensureAltKey();

await ensureConnRef(`${P}_dataverse`, 'FlowOptLab Dataverse', 'shared_commondataserviceforapps');
await ensureConnRef(`${P}_sharepoint`, 'FlowOptLab SharePoint', 'shared_sharepointonline');

const pub = await dv('POST', 'PublishAllXml', {});
console.log(pub.ok ? '✅ published customizations' : `⚠️  PublishAllXml returned ${pub.status}`);
console.log('\nNext: open make.powerapps.com > Solutions > Flow Optimization Lab > Connection references,');
console.log('      bind both references to a connection, then run 02-create-poc-flows.mjs');
