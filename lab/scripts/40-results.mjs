// Prints a Markdown results table from pa_benchmarkrun: for every scenario and volume, the latest
// finished run of each variant (rollback demos with injectBadItem excluded), the speed-up (SLOW ÷ FAST duration) and whether the checksums match.
//
//   node lab/scripts/40-results.mjs [--label cli]  > lab/RESULTS-table.md
import { dvOk, PREFIX } from '../../poc/scripts/lib/dv.mjs';

const P = PREFIX;
const label = process.argv.includes('--label') ? process.argv[process.argv.indexOf('--label') + 1] : undefined;
const cols = ['name', 'scenario', 'variant', 'status', 'volume', 'durationms', 'actionsestimated', 'connectorcalls', 'payloadbytes', 'recordsprocessed', 'resultchecksum', 'starttime', 'label', 'details'].map((c) => `${P}_${c}`);
let url = `${P}_benchmarkruns?$select=${cols.join(',')}&$filter=${P}_scenario ne 'POC' and ${P}_scenario ne 'ADMIN' and ${P}_status ne 'Running'${label ? ` and ${P}_label eq '${label}'` : ''}&$orderby=${P}_starttime desc`;
const rows = [];
while (url) {
  const r = await dvOk('GET', url, undefined, { Prefer: 'odata.maxpagesize=5000' });
  rows.push(...r.json.value);
  url = r.json['@odata.nextLink'];
}
const g = (r, c) => r[`${P}_${c}`];
const groups = new Map();
for (const r of rows) {
  if (g(r, 'scenario') === 'UC7' && g(r, 'variant') !== 'DRIVER') continue;
  if ((g(r, 'details') || '').includes('"injectBadItem":true')) continue; // rollback demos are not benchmarks
  const key = `${g(r, 'scenario')}|${g(r, 'volume')}`;
  if (!groups.has(key)) groups.set(key, {});
  const grp = groups.get(key);
  if (!grp[g(r, 'variant')]) grp[g(r, 'variant')] = r; // newest first
}
const order = (k) => { const [s, v] = k.split('|'); return [Number(s.replace(/\D/g, '')) + (s.includes('DV') ? 0.5 : 0), Number(v)]; };
const keys = [...groups.keys()].sort((a, b) => { const [x1, y1] = order(a); const [x2, y2] = order(b); return x1 - x2 || y1 - y2; });
const fmt = (ms) => (ms == null ? '–' : ms < 120000 ? `${(ms / 1000).toFixed(1)} s` : `${(ms / 60000).toFixed(1)} min`);
const kb = (b) => (!b ? '–' : b < 1048576 ? `${Math.round(b / 1024)} KB` : `${(b / 1048576).toFixed(1)} MB`);
console.log('| UC | Volume | SLOW | MID | FAST | Speed-up | Actions SLOW → FAST | Calls SLOW → FAST | Payload SLOW → FAST | Same result |');
console.log('|---|---:|---:|---:|---:|---:|---|---|---|---|');
for (const k of keys) {
  const [uc, vol] = k.split('|');
  const v = groups.get(k);
  if (!v.SLOW || !v.FAST) continue;
  const s = v.SLOW; const f = v.FAST; const m = v.MID;
  const speed = g(s, 'durationms') && g(f, 'durationms') ? `${(g(s, 'durationms') / g(f, 'durationms')).toFixed(g(s, 'durationms') / g(f, 'durationms') >= 10 ? 0 : 1)}×` : '–';
  const same = g(s, 'resultchecksum') === g(f, 'resultchecksum') ? '✅' : `❌ ${g(s, 'resultchecksum')} vs ${g(f, 'resultchecksum')}`;
  console.log(`| ${uc} | ${Number(vol).toLocaleString('en-US')} | ${fmt(g(s, 'durationms'))} | ${m ? fmt(g(m, 'durationms')) : ''} | ${fmt(g(f, 'durationms'))} | **${speed}** | ${g(s, 'actionsestimated') ?? '–'} → ${g(f, 'actionsestimated') ?? '–'} | ${g(s, 'connectorcalls') ?? '–'} → ${g(f, 'connectorcalls') ?? '–'} | ${kb(g(s, 'payloadbytes'))} → ${kb(g(f, 'payloadbytes'))} | ${same} |`);
}
