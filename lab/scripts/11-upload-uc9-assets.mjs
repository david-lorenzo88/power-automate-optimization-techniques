// UC9 (Excel) assets: uploads lab/assets/Uc9Orders.xlsx (as Uc9Slow.xlsx + Uc9Fast.xlsx) and the Office Scripts in
// lab/assets/office-scripts/*.ts (wrapped as .osts) to <site>/Shared Documents/FlowOptLab,
// then records the drive/file ids that "Run script from SharePoint library" needs in lab.config.json.
//
//   node lab/scripts/11-upload-uc9-assets.mjs [--force-workbook]
//
// Scripts are always overwritten; the workbook only with --force-workbook (it holds demo data).
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { azToken, sp, spHttp, spOk } from './lib/sp.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const assets = join(here, '..', 'assets');
const cfgFile = join(here, '..', 'lab.config.json');
const cfg = JSON.parse(readFileSync(cfgFile, 'utf8'));
const FOLDER = 'FlowOptLab';
const forceWorkbook = process.argv.includes('--force-workbook');

// .osts = JSON envelope around the TypeScript source. parameterInfo describes main()'s
// parameters (after `workbook`) and return type so Power Automate can pass/return values.
function osts(source) {
  const sig = source.match(/function main\(\s*workbook:\s*ExcelScript\.Workbook\s*(?:,([^)]*))?\)\s*:\s*(\w+)/);
  if (!sig) throw new Error('main(workbook: ExcelScript.Workbook, ...) not found');
  const params = (sig[1] || '').split(',').map((s) => s.trim()).filter(Boolean).map((s) => {
    const [name, type] = s.split(':').map((x) => x.trim());
    return { name, type: type === 'number' ? 'number' : type === 'boolean' ? 'boolean' : 'string' };
  });
  const ret = sig[2];
  const parameterInfo = {
    version: 1,
    originalParameterOrder: params.map((p, index) => ({ name: p.name, index })),
    parameterSchema: { type: 'object', required: params.map((p) => p.name), properties: Object.fromEntries(params.map((p) => [p.name, { type: p.type }])) },
    returnSchema: ret === 'void' ? { type: 'object', properties: {} } : { type: 'object', properties: { result: { type: ret === 'number' ? 'number' : ret === 'boolean' ? 'boolean' : 'string' } } },
    signature: { comment: '', parameters: [{ name: 'workbook', comment: '' }, ...params.map((p) => ({ name: p.name, comment: '' }))] },
  };
  return JSON.stringify({ version: '0.3.0', body: source, description: '', noCodeMetadata: '', parameterInfo: JSON.stringify(parameterInfo), apiInfo: JSON.stringify({ variant: 'synchronous', variantVersion: 2 }) });
}

const ctx = await sp();
const libPath = new URL(`${ctx.site}/Shared Documents`).pathname; // server-relative path of the default library
const serverRel = (name) => `${libPath}/${FOLDER}/${name}`;

async function exists(name) {
  return (await spHttp('GET', `_api/web/GetFileByServerRelativePath(decodedurl='${serverRel(name)}')?$select=Name`)).ok;
}

async function upload(name, bytes) {
  if (await exists(name)) {
    await spOk('POST', `_api/web/GetFileByServerRelativePath(decodedurl='${serverRel(name)}')`, { headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' } });
  }
  const url = `${ctx.base}/datasets/${ctx.dataset}/files?folderPath=${encodeURIComponent(`/Shared Documents/${FOLDER}`)}&name=${encodeURIComponent(name)}&queryParametersSingleEncoded=true`;
  const r = await fetch(url, { method: 'POST', headers: { Authorization: `Bearer ${azToken('https://apihub.azure.com')}`, 'Content-Type': 'application/octet-stream' }, body: bytes });
  if (!r.ok) throw new Error(`upload ${name} -> ${r.status}: ${(await r.text()).slice(0, 400)}`);
  console.log(`⬆️  ${name} (${bytes.length} bytes)`);
}

async function driveItem(name) {
  const r = await spOk('GET', `_api/v2.1/drive/root:/${FOLDER}/${encodeURIComponent(name)}?$select=id,name,parentReference`);
  return { id: r.json.id, driveId: r.json.parentReference.driveId };
}

// workbooks: one copy per variant, so SLOW and FAST can race without sharing a table
const workbooks = { SLOW: 'Uc9Slow.xlsx', FAST: 'Uc9Fast.xlsx' };
for (const wbName of Object.values(workbooks)) {
  if (forceWorkbook || !(await exists(wbName))) await upload(wbName, readFileSync(join(assets, 'Uc9Orders.xlsx')));
  else console.log(`ℹ️  ${wbName} exists (use --force-workbook to replace it)`);
}

// scripts
const scripts = {};
for (const f of readdirSync(join(assets, 'office-scripts')).filter((x) => x.endsWith('.ts'))) {
  const name = `${basename(f, '.ts')}.osts`;
  await upload(name, Buffer.from(osts(readFileSync(join(assets, 'office-scripts', f), 'utf8')), 'utf8'));
  scripts[basename(f, '.ts')] = (await driveItem(name)).id;
}

const files = {};
let driveId;
for (const [variant, wbName] of Object.entries(workbooks)) {
  const it = await driveItem(wbName);
  files[variant] = it.id;
  driveId = it.driveId;
}
cfg.uc9 = { source: ctx.site, driveId, files, scriptDriveId: driveId, scripts, folder: `Shared Documents/${FOLDER}` };
writeFileSync(cfgFile, JSON.stringify(cfg, null, 2) + '\n');
console.log('📝 lab.config.json uc9 =', JSON.stringify(cfg.uc9, null, 2));
