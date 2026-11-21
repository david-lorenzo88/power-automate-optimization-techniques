// SharePoint helper for the lab scripts.
//
// The scripts don't call SharePoint directly (the Azure CLI app can't get a
// SharePoint token in many tenants). Instead they go through the SAME
// SharePoint connection the flows use: the connection bound to the `pa_sharepoint`
// connection reference, invoked via the connector runtime ("Send an HTTP request
// to SharePoint"). Tokens come from your `az login`.
//
//   SP_SITE_URL   site that holds the demo lists (default: tenant root site)
import { execSync } from 'node:child_process';
import { dvOk, PREFIX } from '../../../poc/scripts/lib/dv.mjs';

const tokens = {};
export function azToken(resource) {
  const t = tokens[resource];
  if (t && t.exp > Date.now() + 120_000) return t.value;
  const out = JSON.parse(execSync(`az account get-access-token --resource ${resource} -o json`, { encoding: 'utf8' }));
  tokens[resource] = { value: out.accessToken, exp: new Date(out.expiresOn).getTime() };
  return out.accessToken;
}

export async function envId() {
  const r = await dvOk('GET', "RetrieveCurrentOrganization(AccessType=@p1)?@p1=Microsoft.Dynamics.CRM.EndpointAccessType'Default'");
  return r.json.Detail.EnvironmentId;
}

async function powerapps(path) {
  const r = await fetch(`https://api.powerapps.com${path}`, { headers: { Authorization: `Bearer ${azToken('https://service.powerapps.com/')}` } });
  if (!r.ok) throw new Error(`api.powerapps.com ${path} -> ${r.status}: ${await r.text()}`);
  return r.json();
}

/** Runtime base URL for the connection bound to a connection reference. */
export async function connectionRuntime(refLogicalName, apiName) {
  const env = await envId();
  const ref = await dvOk('GET', `connectionreferences?$select=connectionid&$filter=connectionreferencelogicalname eq '${refLogicalName}'`);
  const connectionId = ref.json.value[0]?.connectionid;
  if (!connectionId) throw new Error(`Connection reference ${refLogicalName} is not bound to a connection yet`);
  const filter = `$filter=environment%20eq%20%27${env}%27`;
  const conn = await powerapps(`/providers/Microsoft.PowerApps/apis/${apiName}/connections/${connectionId}?api-version=2016-11-01&${filter}`);
  const bad = (conn.properties.statuses || []).filter((s) => s.status !== 'Connected');
  if (bad.length) throw new Error(`Connection ${connectionId} is not healthy: ${JSON.stringify(bad).slice(0, 400)}`);
  const api = await powerapps(`/providers/Microsoft.PowerApps/apis/${apiName}?api-version=2016-11-01&${filter}`);
  return { env, connectionId, base: `${api.properties.runtimeUrls[0]}${connectionId}` };
}

let spCtx;
export async function sp() {
  if (spCtx) return spCtx;
  const rt = await connectionRuntime(`${PREFIX}_sharepoint`, 'shared_sharepointonline');
  let site = process.env.SP_SITE_URL;
  if (!site) {
    // Tenant root site, as seen by the Graph token of the az CLI.
    const r = await fetch('https://graph.microsoft.com/v1.0/sites/root?$select=webUrl', { headers: { Authorization: `Bearer ${azToken('https://graph.microsoft.com')}` } });
    site = (await r.json()).webUrl;
  }
  site = site.replace(/\/+$/, '');
  const dataset = encodeURIComponent(encodeURIComponent(site));
  spCtx = { ...rt, site, dataset };
  return spCtx;
}

/** Send an HTTP request to SharePoint through the connection. Returns { status, json, text }. */
export async function spHttp(method, uri, { headers = {}, body } = {}) {
  const c = await sp();
  const req = { method, uri, headers: { Accept: 'application/json;odata=nometadata', ...headers } };
  if (body !== undefined) {
    req.body = typeof body === 'string' ? body : JSON.stringify(body);
    if (!req.headers['Content-Type']) req.headers['Content-Type'] = 'application/json;odata=verbose';
  }
  for (let attempt = 0; ; attempt++) {
    const r = await fetch(`${c.base}/datasets/${c.dataset}/httprequest`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${azToken('https://apihub.azure.com')}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(req),
    });
    const text = await r.text();
    if ((r.status === 429 || r.status >= 502) && attempt < 4) {
      await new Promise((res) => setTimeout(res, 2000 * 2 ** attempt));
      continue;
    }
    let json;
    try { json = JSON.parse(text); } catch { json = undefined; }
    return { status: r.status, ok: r.ok, json, text };
  }
}

export async function spOk(method, uri, opts) {
  const r = await spHttp(method, uri, opts);
  if (!r.ok) throw new Error(`SP ${method} ${uri} -> ${r.status}: ${r.text.slice(0, 600)}`);
  return r;
}

const odataType = (kind) => ({ __metadata: { type: kind } });

export async function ensureList(title, description) {
  const found = await spHttp('GET', `_api/web/lists/getbytitle('${title}')?$select=Id,ItemCount,ListItemEntityTypeFullName`);
  if (found.ok) return { ...found.json, created: false };
  await spOk('POST', '_api/web/lists', {
    body: { ...odataType('SP.List'), BaseTemplate: 100, Title: title, Description: description, AllowContentTypes: true },
  });
  const r = await spOk('GET', `_api/web/lists/getbytitle('${title}')?$select=Id,ItemCount,ListItemEntityTypeFullName`);
  return { ...r.json, created: true };
}

/** fields: [{ name, kind: 'text'|'number'|'date'|'lookup', indexed?, lookupListId? }] */
export async function ensureFields(listTitle, fields) {
  const have = new Set((await spOk('GET', `_api/web/lists/getbytitle('${listTitle}')/fields?$select=InternalName&$top=500`)).json.value.map((f) => f.InternalName));
  const added = [];
  for (const f of fields) {
    if (!have.has(f.name)) {
      if (f.kind === 'lookup') {
        await spOk('POST', `_api/web/lists/getbytitle('${listTitle}')/fields/addfield`, {
          body: { parameters: { ...odataType('SP.FieldCreationInformation'), FieldTypeKind: 7, Title: f.name, LookupListId: f.lookupListId, LookupFieldName: 'Title' } },
        });
      } else {
        const kind = { text: [2, 'SP.Field'], number: [9, 'SP.FieldNumber'], date: [4, 'SP.FieldDateTime'] }[f.kind];
        await spOk('POST', `_api/web/lists/getbytitle('${listTitle}')/fields`, {
          body: { ...odataType(kind[1]), Title: f.name, FieldTypeKind: kind[0], ...(f.kind === 'date' ? { DisplayFormat: 0 } : {}) },
        });
      }
      added.push(f.name);
    }
    if (f.indexed) {
      await spOk('POST', `_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${f.name}')`, {
        headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
        body: { ...odataType('SP.Field'), Indexed: true },
      });
    }
  }
  return added;
}
