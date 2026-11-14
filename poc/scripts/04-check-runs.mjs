// Checklist #9 (+ helps #6): compare what the flows reported in pa_benchmarkrun
// with what the platform stored in the flowrun table, and print workflow() tags.
//
//   node poc/scripts/04-check-runs.mjs [top=10]
import { dv, dvOk, PREFIX, table } from './lib/dv.mjs';

const P = PREFIX;
const top = Number(process.argv[2] || 10);

const runs = await dvOk('GET',
  `${P}_benchmarkruns?$select=${P}_name,${P}_variant,${P}_status,${P}_durationms,${P}_runid,${P}_flowid,${P}_details,createdon&$orderby=createdon desc&$top=${top}`);

const rows = [];
let tagsShown = false;
for (const r of runs.json.value) {
  const runId = r[`${P}_runid`];
  let fr;
  const q = await dv('GET', `flowruns?$filter=name eq '${runId}'&$top=1`);
  if (q.ok) fr = q.json.value[0];
  else if (!rows.length) console.log(`⚠️  flowrun query failed (${q.status}): ${q.json?.error?.message ?? q.text}`);

  if (!tagsShown && r[`${P}_details`]?.startsWith('{"')) {
    try {
      const wf = JSON.parse(r[`${P}_details`]);
      if (wf.tags) {
        console.log('workflow() tags exposed by the platform:', wf.tags, '\n');
        tagsShown = true;
      }
    } catch { /* details is not workflow() json for this row */ }
  }

  rows.push({
    created: r.createdon,
    variant: r[`${P}_variant`],
    status: r[`${P}_status`],
    flowMs: r[`${P}_durationms`],
    flowrunDuration: fr?.duration ?? '(not found)',
    flowrunStatus: fr?.status ?? '',
    runId,
  });
}
table(rows);

if (rows.some((x) => x.flowrunDuration !== '(not found)')) {
  console.log('\nℹ️  Compare flowMs with flowrunDuration to confirm the unit of flowrun.duration (docs disagree: seconds vs ms).');
} else {
  const sample = await dv('GET', 'flowruns?$top=1&$orderby=createdon desc');
  if (sample.ok && sample.json.value[0]) {
    console.log('\nNo match by name. Latest flowrun row, to see which column holds the run id:');
    console.log(Object.fromEntries(Object.entries(sample.json.value[0]).filter(([k]) => !k.startsWith('@'))));
  }
  console.log('flowrun rows can take a few minutes to appear after a run finishes.');
}
