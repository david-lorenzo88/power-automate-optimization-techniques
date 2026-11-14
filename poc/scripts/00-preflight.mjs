// Checklist #1 (partial) + #9 (partial): can we talk to Dataverse, who are we,
// is the flowrun table readable, is the publisher prefix free?
import { dv, dvOk, ORG, PREFIX, SOLUTION } from './lib/dv.mjs';

console.log(`Org: ${ORG}\nSolution: ${SOLUTION}\nPrefix: ${PREFIX}\n`);

const who = await dvOk('GET', 'WhoAmI');
console.log('✅ WhoAmI', { userId: who.json.UserId, orgId: who.json.OrganizationId, ms: who.ms });

const user = await dv('GET', `systemusers(${who.json.UserId})?$select=fullname,domainname`);
if (user.ok) console.log('   user:', user.json.fullname, user.json.domainname);

const fr = await dv('GET', 'flowruns?$top=1');
if (fr.ok) {
  const sample = fr.json.value?.[0];
  console.log(`✅ flowrun table readable (${fr.json.value.length} sample row)`);
  if (sample) console.log('   flowrun columns:', Object.keys(sample).filter((k) => !k.startsWith('@')).join(', '));
} else {
  console.log(`❌ flowrun table not readable (${fr.status}). Check "Cloud flow run history in Dataverse" and your security role.`);
}

const pubs = await dvOk('GET', `publishers?$select=uniquename,customizationprefix&$filter=customizationprefix eq '${PREFIX}'`);
if (pubs.json.value.length) {
  console.log(`ℹ️  prefix '${PREFIX}' already used by publisher(s):`, pubs.json.value.map((p) => p.uniquename).join(', '));
  console.log('   That is fine if it is our publisher (flowoptlab); otherwise set PUBLISHER_PREFIX to something else.');
} else {
  console.log(`✅ prefix '${PREFIX}' is free`);
}

const sol = await dvOk('GET', `solutions?$select=uniquename,version&$filter=uniquename eq '${SOLUTION}'`);
console.log(sol.json.value.length ? `ℹ️  solution ${SOLUTION} already exists (v${sol.json.value[0].version})` : `ℹ️  solution ${SOLUTION} not created yet → run 01-setup-dataverse.mjs`);

console.log('\nManual checks for checklist #1:');
console.log('  - PPAC > Environments > <env> > Settings > Product > Features > "Enable code apps" = On');
console.log('  - Your user has Power Apps Premium (or the environment is pay-as-you-go)');
