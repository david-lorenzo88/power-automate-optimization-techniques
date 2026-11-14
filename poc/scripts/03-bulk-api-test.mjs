// Checklist #7 (API side): measures Dataverse single requests vs bulk messages vs $batch,
// and proves changeset rollback. Uses the pa_orderimport table from 01-setup-dataverse.mjs.
//
//   node poc/scripts/03-bulk-api-test.mjs [N=1000]
//
// Note: these are timings from YOUR machine to Dataverse (client side), not from a flow.
// They tell us what the platform can do; the flows will add their own overhead.
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { dv, dvOk, API, PREFIX, uuid, table } from './lib/dv.mjs';

const P = PREFIX;
const N = Number(process.argv[2] || 1000);
const SINGLE_N = Math.min(N, 100);
const SET = `${P}_orderimports`;
const TYPE = `Microsoft.Dynamics.CRM.${P}_orderimport`;
const tag = `poc-${new Date().toISOString().replace(/[:.]/g, '-')}`;
const results = [];
const regions = ['EMEA', 'AMER', 'APAC'];

const rec = (i, extra = {}) => ({
  [`${P}_name`]: `Order ${i}`,
  [`${P}_externalid`]: `${tag}-${i}`,
  [`${P}_region`]: regions[i % 3],
  [`${P}_amount`]: (i % 97) * 10.5,
  [`${P}_batchtag`]: tag,
  ...extra,
});
const chunk = (arr, size) => Array.from({ length: Math.ceil(arr.length / size) }, (_, i) => arr.slice(i * size, i * size + size));
const add = (test, records, requests, ms, note = '') =>
  results.push({ test, records, httpRequests: requests, ms, msPerRecord: records ? +(ms / records).toFixed(1) : 0, note });

async function countTag(t = tag) {
  const r = await dvOk('GET', `${SET}?$filter=${P}_batchtag eq '${t}'&$count=true&$top=1&$select=${P}_orderimportid`);
  return r.json['@odata.count'];
}
async function idsByTag(t = tag) {
  let url = `${SET}?$filter=${P}_batchtag eq '${t}'&$select=${P}_orderimportid`;
  const ids = [];
  while (url) {
    const r = await dvOk('GET', url, undefined, { Prefer: 'odata.maxpagesize=5000' });
    ids.push(...r.json.value.map((x) => x[`${P}_orderimportid`]));
    url = r.json['@odata.nextLink'];
  }
  return ids;
}

function batchBody(ops, { changeset }) {
  const b = `batch_${uuid()}`;
  const cs = `changeset_${uuid()}`;
  const part = (op, i) => [
    `--${changeset ? cs : b}`,
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    ...(changeset ? [`Content-ID: ${i + 1}`] : []),
    '',
    `${op.method} ${API}/${op.path} HTTP/1.1`,
    ...(op.body ? ['Content-Type: application/json; type=entry', '', JSON.stringify(op.body)] : ['', '']),
  ].join('\r\n');
  const lines = changeset
    ? [`--${b}`, `Content-Type: multipart/mixed; boundary=${cs}`, '', ops.map(part).join('\r\n'), `--${cs}--`, `--${b}--`, '']
    : [ops.map(part).join('\r\n'), `--${b}--`, ''];
  return { body: lines.join('\r\n'), boundary: b };
}
async function sendBatch(ops, opts) {
  const { body, boundary } = batchBody(ops, opts);
  const headers = { 'Content-Type': `multipart/mixed; boundary=${boundary}` };
  if (opts.continueOnError) headers.Prefer = 'odata.continue-on-error';
  return dv('POST', '$batch', body, headers);
}

console.log(`Tag: ${tag}   N=${N}\n`);

// 1) Baseline: one POST per record (sequential) = what "Add a new row" in Apply to each does
{
  const t0 = performance.now();
  for (let i = 0; i < SINGLE_N; i++) await dvOk('POST', SET, rec(i));
  const ms = Math.round(performance.now() - t0);
  add('Single POST (sequential)', SINGLE_N, SINGLE_N, ms, SINGLE_N < N ? `extrapolated to N=${N}: ~${Math.round((ms / SINGLE_N) * N / 1000)} s` : '');
}

// 2) CreateMultiple in chunks of 1000
const createItems = Array.from({ length: N }, (_, i) => ({ ...rec(SINGLE_N + i), '@odata.type': TYPE }));
{
  const t0 = performance.now();
  const chunks = chunk(createItems, 1000);
  for (const c of chunks) await dvOk('POST', `${SET}/Microsoft.Dynamics.CRM.CreateMultiple`, { Targets: c });
  add('CreateMultiple (1000/request)', N, chunks.length, Math.round(performance.now() - t0));
}

// 3) UpdateMultiple on everything created so far
{
  const ids = await idsByTag();
  const t0 = performance.now();
  const chunks = chunk(ids, 1000);
  for (const c of chunks) {
    await dvOk('POST', `${SET}/Microsoft.Dynamics.CRM.UpdateMultiple`, {
      Targets: c.map((id) => ({ '@odata.type': TYPE, [`${P}_orderimportid`]: id, [`${P}_region`]: 'UPDATED' })),
    });
  }
  add('UpdateMultiple (1000/request)', ids.length, chunks.length, Math.round(performance.now() - t0));
}

// 4) UpsertMultiple via alternate key: half existing, half new
{
  const half = Math.floor(N / 2);
  const targets = [
    ...Array.from({ length: half }, (_, i) => SINGLE_N + i), // existing
    ...Array.from({ length: N - half }, (_, i) => SINGLE_N + N + i), // new
  ].map((i) => {
    const { [`${P}_externalid`]: key, ...rest } = rec(i, { [`${P}_region`]: 'UPSERTED' });
    return { '@odata.type': TYPE, '@odata.id': `${SET}(${P}_externalid='${key}')`, ...rest };
  });
  const t0 = performance.now();
  const r = await dv('POST', `${SET}/Microsoft.Dynamics.CRM.UpsertMultiple`, { Targets: targets });
  add('UpsertMultiple (alt key, 50% new)', N, 1, Math.round(performance.now() - t0),
    r.ok ? '' : `FAILED ${r.status}: ${r.json?.error?.message ?? ''} (alternate key index may still be building)`);
}

// 5) $batch, one changeset of 100 creates, all valid → expect 100 rows
{
  const t = `${tag}-cs-ok`;
  const ops = Array.from({ length: 100 }, (_, i) => ({ method: 'POST', path: SET, body: rec(i, { [`${P}_batchtag`]: t, [`${P}_externalid`]: `${t}-${i}` }) }));
  const r = await sendBatch(ops, { changeset: true });
  const created = await countTag(t);
  add('$batch changeset 100 creates (valid)', 100, 1, r.ms, `HTTP ${r.status}, rows created: ${created}`);
}

// 6) $batch, one changeset of 100 creates with ONE bad record → expect 0 rows (rollback)
{
  const t = `${tag}-cs-bad`;
  const ops = Array.from({ length: 100 }, (_, i) => ({
    method: 'POST', path: SET,
    body: rec(i, { [`${P}_batchtag`]: t, [`${P}_externalid`]: `${t}-${i}`, ...(i === 50 ? { [`${P}_amount`]: 'not-a-number' } : {}) }),
  }));
  const r = await sendBatch(ops, { changeset: true });
  const created = await countTag(t);
  add('$batch changeset 100 creates (1 invalid)', 100, 1, r.ms, `HTTP ${r.status}, rows created: ${created} → ${created === 0 ? '✅ rolled back' : '❌ NOT atomic'}`);
}

// 7) $batch without changeset + continue-on-error, 1 bad record → expect 99 rows
{
  const t = `${tag}-nocs-bad`;
  const ops = Array.from({ length: 100 }, (_, i) => ({
    method: 'POST', path: SET,
    body: rec(i, { [`${P}_batchtag`]: t, [`${P}_externalid`]: `${t}-${i}`, ...(i === 50 ? { [`${P}_amount`]: 'not-a-number' } : {}) }),
  }));
  const r = await sendBatch(ops, { changeset: false, continueOnError: true });
  const created = await countTag(t);
  add('$batch no changeset + continue-on-error (1 invalid)', 100, 1, r.ms, `HTTP ${r.status}, rows created: ${created} (expect 99)`);
}

// 8) Cleanup: $batch DELETE, 1000 per request, continue-on-error
{
  const all = [];
  for (const t of [tag, `${tag}-cs-ok`, `${tag}-cs-bad`, `${tag}-nocs-bad`]) all.push(...(await idsByTag(t)));
  const t0 = performance.now();
  const chunks = chunk(all, 1000);
  for (const c of chunks) await sendBatch(c.map((id) => ({ method: 'DELETE', path: `${SET}(${id})` })), { changeset: false, continueOnError: true });
  add('$batch DELETE (1000/request)', all.length, chunks.length, Math.round(performance.now() - t0), `remaining: ${await countTag()}`);
}

table(results);
const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'results');
mkdirSync(outDir, { recursive: true });
const file = join(outDir, `bulk-api-${tag}.json`);
writeFileSync(file, JSON.stringify({ tag, N, results }, null, 2));
console.log(`\nSaved ${file}`);
