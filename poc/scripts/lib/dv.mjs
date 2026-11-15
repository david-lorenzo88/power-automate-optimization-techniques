// Minimal Dataverse Web API helper (Node 18+, no dependencies).
//
// Configuration (environment variables):
//   DATAVERSE_URL    https://yourorg.crm4.dynamics.com   (required)
//   DATAVERSE_TOKEN  bearer token (optional). If not set, the helper calls
//                    `az account get-access-token --resource <DATAVERSE_URL>`
//                    so run `az login` first (Azure CLI).
//   SOLUTION_NAME    defaults to FlowOptimizationLab
//   PUBLISHER_PREFIX defaults to pa

import { execSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';

export const ORG = (process.env.DATAVERSE_URL || '').replace(/\/+$/, '');
export const SOLUTION = process.env.SOLUTION_NAME || 'FlowOptimizationLab';
export const PREFIX = process.env.PUBLISHER_PREFIX || 'pa';
export const API = `${ORG}/api/data/v9.2`;

if (!ORG) {
  console.error('Set DATAVERSE_URL, e.g.  export DATAVERSE_URL=https://yourorg.crm4.dynamics.com');
  process.exit(1);
}

let cachedToken;
export function token() {
  if (cachedToken) return cachedToken;
  if (process.env.DATAVERSE_TOKEN) return (cachedToken = process.env.DATAVERSE_TOKEN);
  try {
    cachedToken = execSync(`az account get-access-token --resource ${ORG} --query accessToken -o tsv`, {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }).trim();
    return cachedToken;
  } catch (e) {
    console.error('Could not get a token with Azure CLI. Run `az login` (or set DATAVERSE_TOKEN).');
    console.error(String(e.stderr || e.message));
    process.exit(1);
  }
}

export async function dv(method, path, body, extraHeaders = {}) {
  const url = path.startsWith('http') ? path : `${API}/${path.replace(/^\//, '')}`;
  const headers = {
    Authorization: `Bearer ${token()}`,
    Accept: 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    ...extraHeaders,
  };
  let payload;
  if (body !== undefined) {
    if (typeof body === 'string') {
      payload = body;
    } else {
      headers['Content-Type'] = 'application/json; charset=utf-8';
      payload = JSON.stringify(body);
    }
  }
  const started = performance.now();
  // A stalled connection must not hang a script forever: 2-minute timeout, retry network errors twice.
  let res;
  for (let attempt = 0; ; attempt++) {
    try {
      res = await fetch(url, { method, headers, body: payload, signal: AbortSignal.timeout(120_000) });
      break;
    } catch (e) {
      if (attempt >= 2) throw e;
      await new Promise((r) => setTimeout(r, 3000 * (attempt + 1)));
    }
  }
  const ms = Math.round(performance.now() - started);
  const text = await res.text();
  let json;
  try { json = text ? JSON.parse(text) : undefined; } catch { json = undefined; }
  return { status: res.status, ok: res.ok, ms, json, text, headers: res.headers };
}

export async function dvOk(method, path, body, extraHeaders) {
  const r = await dv(method, path, body, extraHeaders);
  if (!r.ok) {
    const msg = r.json?.error?.message || r.text;
    throw new Error(`${method} ${path} -> ${r.status}: ${msg}`);
  }
  return r;
}

export const label = (text) => ({
  '@odata.type': 'Microsoft.Dynamics.CRM.Label',
  LocalizedLabels: [{ '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 }],
});

export const solutionHeader = () => ({ 'MSCRM.SolutionUniqueName': SOLUTION });

export const uuid = () => randomUUID();

export function table(rows) {
  if (!rows.length) return console.log('(no rows)');
  console.table(rows);
}
