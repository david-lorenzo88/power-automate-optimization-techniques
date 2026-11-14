import { useEffect, useMemo, useState } from 'react';
import { getContext } from '@microsoft/power-apps/app';
import type { IContext } from '@microsoft/power-apps/app';
import { discoverServices } from './lib/services';
import type { DiscoveredService } from './lib/services';
import { launchAndTrack } from './lib/benchmark';
import type { TrackedRun } from './lib/benchmark';
import './App.css';

type Verdict = 'pass' | 'fail' | 'info' | 'running';

interface TestDef {
  id: string;
  checklist: string;
  title: string;
  flow: string;
  expectation: string;
  payload: () => Record<string, unknown>;
  evaluate: (r: TrackedRun) => { verdict: Verdict; note: string };
}

const secs = (ms?: number) => (ms == null ? '–' : `${(ms / 1000).toFixed(1)} s`);
const rowOk = (r: TrackedRun) => r.row?.status === 'Succeeded';

function buildTests(sp: { siteUrl: string; listName: string; batchSize: number; injectBadItem: boolean }): TestDef[] {
  return [
    {
      id: 'T3', checklist: '#3', title: 'Invoke a flow from the Code App', flow: 'PocRespondLate',
      expectation: 'Run() resolves with success = true and returns runId. Telemetry row ends as Succeeded.',
      payload: () => ({ delaySeconds: 1, label: 'poc-t3' }),
      evaluate: (r) => ({
        verdict: r.invoke.invokeSuccess && rowOk(r) ? 'pass' : 'fail',
        note: `Run() ${r.invoke.invokeSuccess ? 'ok' : 'failed'} in ${secs(r.invoke.invokeMs)}; response: ${JSON.stringify(r.invoke.responseData ?? null)}`,
      }),
    },
    {
      id: 'T4', checklist: '#4', title: 'Respond early, keep working (60 s)', flow: 'PocRespondEarly',
      expectation: 'Run() resolves within a few seconds. The flow keeps running and the row becomes Succeeded after about 60 s.',
      payload: () => ({ delaySeconds: 60, label: 'poc-t4' }),
      evaluate: (r) => {
        const early = (r.invoke.invokeMs ?? Infinity) < 30_000;
        const kept = rowOk(r) && (r.row?.durationMs ?? 0) >= 55_000;
        return {
          verdict: early && kept ? 'pass' : 'fail',
          note: `Run() resolved after ${secs(r.invoke.invokeMs)} (${early ? 'early ✅' : 'NOT early ❌'}); flow duration ${secs(r.row?.durationMs)} (${kept ? 'kept running ✅' : '❌'})`,
        };
      },
    },
    {
      id: 'T5a', checklist: '#5', title: '> 120 s, no Respond action (150 s)', flow: 'PocNoRespond',
      expectation: 'Fire-and-forget: the app never breaks and the row becomes Succeeded after about 150 s.',
      payload: () => ({ delaySeconds: 150, label: 'poc-t5a' }),
      evaluate: (r) => ({
        verdict: rowOk(r) ? 'pass' : 'fail',
        note: `Run() ${r.invoke.invokeSuccess ? 'ok' : `error: ${r.invoke.invokeError}`} after ${secs(r.invoke.invokeMs)}; flow ${r.row?.status} in ${secs(r.row?.durationMs)}`,
      }),
    },
    {
      id: 'T5b', checklist: '#5', title: '> 120 s, Respond at the end (150 s)', flow: 'PocRespondLate',
      expectation: 'Run() probably errors or times out at about 120 s, but the flow still completes. Records how the client behaves.',
      payload: () => ({ delaySeconds: 150, label: 'poc-t5b' }),
      evaluate: (r) => ({
        verdict: rowOk(r) ? (r.invoke.invokeSuccess ? 'pass' : 'info') : 'fail',
        note: `Run() ${r.invoke.invokeSuccess ? 'ok' : `error: ${r.invoke.invokeError}`} after ${secs(r.invoke.invokeMs)}; flow ${r.row?.status} in ${secs(r.row?.durationMs)}`,
      }),
    },
    {
      id: 'T8', checklist: '#8', title: `SharePoint $batch (${sp.batchSize} items${sp.injectBadItem ? ', 1 invalid' : ''})`, flow: 'PocSpBatch',
      expectation: sp.injectBadItem
        ? 'Shows whether a SharePoint changeset is atomic: are the 1 invalid + N-1 valid items all rejected, or only the invalid one?'
        : 'All items are created in one HTTP request. Try 100 / 500 / 1000 / 2000 to find the limit.',
      payload: () => ({ ...sp, label: `poc-t8-${sp.batchSize}${sp.injectBadItem ? '-bad' : ''}` }),
      evaluate: (r) => {
        let d: Record<string, unknown> = {};
        try { d = JSON.parse(r.row?.details ?? '{}'); } catch { /* not json */ }
        return {
          verdict: r.row?.status === 'Succeeded' ? 'pass' : 'info',
          note: `HTTP ${String(d.batchHttpStatus ?? '?')}, created ${String(d.created201 ?? r.row?.recordsProcessed ?? '?')}/${sp.batchSize}, 4xx ${String(d.failed4xx ?? '?')}, batch call ${String(d.batchMs ?? '?')} ms, payload ${r.row?.payloadBytes ?? '?'} bytes`,
        };
      },
    },
  ];
}

interface Result { test: TestDef; run: TrackedRun; verdict: Verdict; note: string }
type LinkMark = Record<string, 'ok' | 'broken' | undefined>;

export default function PocPage() {
  const [ctx, setCtx] = useState<IContext>();
  const [ctxError, setCtxError] = useState<string>();
  const [services, setServices] = useState<DiscoveredService[]>([]);
  const [results, setResults] = useState<Record<string, Result>>({});
  const [busy, setBusy] = useState<Record<string, boolean>>({});
  const [linkMarks, setLinkMarks] = useState<LinkMark>({});
  const [sp, setSp] = useState({ siteUrl: '', listName: 'PocImport', batchSize: 100, injectBadItem: false });

  useEffect(() => {
    const t = setTimeout(() => setCtxError('No Power Apps host context after 8 s. Open the app through the Local Play URL from `npm run dev`, or the published app.'), 8000);
    getContext()
      .then((c) => { clearTimeout(t); setCtx(c); setCtxError(undefined); })
      .catch((e: unknown) => { clearTimeout(t); setCtxError(e instanceof Error ? e.message : String(e)); });
    discoverServices().then(setServices);
  }, []);

  const tests = useMemo(() => buildTests(sp), [sp]);

  async function runTest(t: TestDef) {
    const key = t.id === 'T8' ? `T8-${sp.batchSize}${sp.injectBadItem ? '-bad' : ''}` : t.id;
    setBusy((b) => ({ ...b, [key]: true }));
    const update = (run: TrackedRun) => {
      const ev = run.finished ? t.evaluate(run) : { verdict: 'running' as Verdict, note: `elapsed ${secs(Date.now() - run.startedAt)} · row: ${run.row?.status ?? '(not yet)'} · Run(): ${run.invoke.invokeMs == null ? 'pending' : secs(run.invoke.invokeMs)}` };
      setResults((r) => ({ ...r, [key]: { test: { ...t, id: key }, run, ...ev } }));
    };
    try {
      await launchAndTrack(t.flow, t.payload(), update);
    } catch (e) {
      setResults((r) => ({
        ...r,
        [key]: { test: { ...t, id: key }, run: { correlationId: '-', flow: t.flow, payload: {}, startedAt: Date.now(), invoke: {}, finished: true }, verdict: 'fail', note: e instanceof Error ? e.message : String(e) },
      }));
    } finally {
      setBusy((b) => ({ ...b, [key]: false }));
    }
  }

  const finishedWithRow = Object.values(results).filter((r) => r.run.row?.runId);
  const linkRun = finishedWithRow[finishedWithRow.length - 1]?.run.row;
  const links = linkRun
    ? [
        { key: 'runUrl', label: 'Run (flows path)', url: linkRun.runUrl },
        { key: 'runUrlSolution', label: 'Run (solutions/~preferred path)', url: linkRun.runUrlSolution },
        { key: 'editUrl', label: 'Edit flow (/edit)', url: linkRun.editUrl },
        { key: 'detailsUrl', label: 'Flow details (/details)', url: linkRun.editUrl?.replace(/\/edit$/, '/details') },
        {
          key: 'appBuiltRun', label: 'Run (built in app from getContext)',
          url: ctx && linkRun.flowId ? `https://make.powerautomate.com/environments/${ctx.app.environmentId}/flows/${linkRun.flowId}/runs/${linkRun.runId}` : undefined,
        },
      ]
    : [];

  function markdown(): string {
    const lines = [
      `## PoC results (${new Date().toISOString()})`,
      '',
      `Environment: ${ctx?.app.environmentId ?? '?'} · App URL: ${ctx?.app.appUrl ?? '?'}`,
      '',
      '| Test | Checklist | Verdict | Notes | Correlation |',
      '|---|---|---|---|---|',
      ...Object.values(results).map((r) => `| ${r.test.id} ${r.test.title} | ${r.test.checklist} | ${r.verdict} | ${r.note.replace(/\|/g, '\\|')} | ${r.run.correlationId} |`),
      '',
      '### Links (#6)',
      '',
      ...links.map((l) => `- ${l.label}: ${linkMarks[l.key] ?? 'not tested'} · \`${l.url ?? 'n/a'}\``),
      '',
      '### Generated services',
      '',
      ...services.map((s) => `- ${s.kind}: ${s.exportName} (${s.file})`),
    ];
    return lines.join('\n');
  }

  const [md, setMd] = useState<string>();

  return (
    <div className="poc">
      <header>
        <h1>Flow Optimization Lab · PoC</h1>
        <p className="muted">Checklist tests for the Code App benchmark harness (docs/03-code-app-benchmark-harness.md)</p>
      </header>

      <section className="card">
        <h2>#1 Context</h2>
        {ctx ? (
          <dl className="kv">
            <dt>Environment</dt><dd>{ctx.app.environmentId}</dd>
            <dt>User</dt><dd>{ctx.user.fullName} ({ctx.user.userPrincipalName})</dd>
            <dt>Dataverse</dt><dd>{ctx.app.dataverseOrgUrl ?? '–'}</dd>
            <dt>App URL</dt><dd>{ctx.app.appUrl ?? '–'}</dd>
          </dl>
        ) : (
          <p className={ctxError ? 'bad' : 'muted'}>{ctxError ?? 'Loading context…'}</p>
        )}
      </section>

      <section className="card">
        <h2>#3 Generated services</h2>
        {services.length === 0 ? (
          <p className="bad">Nothing generated yet. Run <code>pa app add data-source --connector dataverse --table pa_benchmarkrun</code> and <code>pa app add flow --flow-id …</code>, then rebuild.</p>
        ) : (
          <ul className="svc">
            {services.filter((s) => s.kind !== 'other').map((s) => (
              <li key={s.file + s.exportName}><span className={`tag ${s.kind}`}>{s.kind}</span> {s.exportName}</li>
            ))}
          </ul>
        )}
      </section>

      <section className="card">
        <h2>Tests</h2>
        <div className="sp-form">
          <label>SharePoint site URL<input value={sp.siteUrl} placeholder="https://contoso.sharepoint.com/sites/FlowLab" onChange={(e) => setSp({ ...sp, siteUrl: e.target.value.trim().replace(/\/+$/, '') })} /></label>
          <label>List<input value={sp.listName} onChange={(e) => setSp({ ...sp, listName: e.target.value })} /></label>
          <label>Batch size
            <select value={sp.batchSize} onChange={(e) => setSp({ ...sp, batchSize: Number(e.target.value) })}>
              {[100, 500, 1000, 2000].map((n) => <option key={n} value={n}>{n}</option>)}
            </select>
          </label>
          <label className="check"><input type="checkbox" checked={sp.injectBadItem} onChange={(e) => setSp({ ...sp, injectBadItem: e.target.checked })} /> inject 1 invalid item</label>
        </div>

        <div className="tests">
          {tests.map((t) => {
            const key = t.id === 'T8' ? `T8-${sp.batchSize}${sp.injectBadItem ? '-bad' : ''}` : t.id;
            const r = results[key];
            const disabled = busy[key] || (t.id === 'T8' && !sp.siteUrl);
            return (
              <article key={t.id} className="test">
                <div className="test-head">
                  <strong>{t.id}</strong> <span className="muted">{t.checklist}</span> {t.title}
                  <button disabled={disabled} onClick={() => runTest(t)}>{busy[key] ? 'Running…' : 'Run'}</button>
                </div>
                <p className="muted small">{t.expectation}</p>
                {r && (
                  <div className={`result ${r.verdict}`}>
                    <span className="pill">{r.verdict}</span> {r.note}
                    {r.run.row?.runUrl && <a href={r.run.row.runUrl} target="_blank" rel="noreferrer"> · open run ↗</a>}
                  </div>
                )}
              </article>
            );
          })}
        </div>
      </section>

      <section className="card">
        <h2>#6 Links</h2>
        {!linkRun ? (
          <p className="muted">Run a test first; links come from its telemetry row.</p>
        ) : (
          <table className="links">
            <tbody>
              {links.map((l) => (
                <tr key={l.key}>
                  <td>{l.label}</td>
                  <td>
                    {l.url ? (
                      <>
                        <a href={l.url} target="_blank" rel="noreferrer">anchor ↗</a>{' '}
                        <button className="link" onClick={() => window.open(l.url, '_blank', 'noopener')}>window.open ↗</button>
                      </>
                    ) : 'n/a'}
                  </td>
                  <td>
                    <button className={linkMarks[l.key] === 'ok' ? 'on' : ''} onClick={() => setLinkMarks({ ...linkMarks, [l.key]: 'ok' })}>works</button>
                    <button className={linkMarks[l.key] === 'broken' ? 'on bad' : ''} onClick={() => setLinkMarks({ ...linkMarks, [l.key]: 'broken' })}>broken</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <section className="card">
        <h2>Export</h2>
        <button onClick={() => { const m = markdown(); setMd(m); navigator.clipboard?.writeText(m).catch(() => undefined); }}>Copy results as Markdown</button>
        {md && <textarea readOnly value={md} rows={14} />}
        <p className="muted small">Paste into <code>poc/RESULTS.md</code> (or send it to Claude).</p>
      </section>
    </div>
  );
}
