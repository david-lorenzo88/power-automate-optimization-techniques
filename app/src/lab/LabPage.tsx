import { useCallback, useEffect, useMemo, useState } from 'react';
import { getHistory, getTriggeredRuns, launchAndTrack } from '../lib/benchmark';
import type { RunRow, TrackedRun } from '../lib/benchmark';
import { hasFlow } from '../lib/services';
import { ADMIN_FLOWS, SCENARIOS } from './scenarios';
import type { Scenario, Variant } from './scenarios';
import { baseline, compare, fmtKb, fmtMs, fmtNum, isDone } from './compare';

const VARIANTS: Variant[] = ['SLOW', 'MID', 'FAST'];
const LONG = 3 * 60 * 60_000; // SLOW runs at full volume can take hours

interface LiveRun extends TrackedRun { scenario: string; variant: string; volume: number }
type Live = Record<string, LiveRun>; // key: `${scenario}:${variant}`

function RunLinks({ row }: { row?: RunRow }) {
  if (!row?.runUrl) return null;
  return (
    <span className="run-links">
      <a href={row.runUrl} target="_blank" rel="noreferrer">open run ↗</a>
      {row.editUrl && <> · <a href={row.editUrl} target="_blank" rel="noreferrer">flow ↗</a></>}
    </span>
  );
}

function Tile({ variant, technique, live, last }: { variant: Variant; technique: string; live?: LiveRun; last?: RunRow }) {
  const row = live?.row ?? last;
  const running = live && !live.finished;
  const status = running ? (row?.status ?? 'Starting') : row?.status;
  const elapsed = running ? Date.now() - live.startedAt : row?.durationMs;
  return (
    <div className={`tile v-${variant.toLowerCase()}`}>
      <div className="tile-head">
        <strong>{variant}</strong>
        {status && <span className={`status s-${status.toLowerCase()}`}>{status}</span>}
        {!live && last && <span className="muted small">last run · vol {fmtNum(last.volume)}</span>}
      </div>
      <p className="technique">{technique}</p>
      {row || running ? (
        <dl className="kpis">
          <dt>Duration</dt><dd className="big">{fmtMs(elapsed)}{running ? '…' : ''}</dd>
          <dt>Actions (est.)</dt><dd>{fmtNum(row?.actionsEstimated)}</dd>
          <dt>Connector calls</dt><dd>{fmtNum(row?.connectorCalls)}</dd>
          <dt>Records</dt><dd>{fmtNum(row?.recordsProcessed)}{row?.recordsRetrieved ? <span className="muted"> / {fmtNum(row.recordsRetrieved)} read</span> : null}</dd>
          <dt>Loop iterations</dt><dd>{fmtNum(row?.loopIterations)}</dd>
          <dt>Payload</dt><dd>{fmtKb(row?.payloadBytes)}</dd>
          <dt>Checksum</dt><dd className="mono">{row?.checksum ?? '–'}</dd>
        </dl>
      ) : <p className="muted small">Not run yet.</p>}
      {live?.invoke.invokeError && <p className="bad small">Run(): {live.invoke.invokeError}</p>}
      {row?.error && <details className="small"><summary className="bad">error</summary><pre>{row.error}</pre></details>}
      <RunLinks row={row} />
    </div>
  );
}

function Bars({ items, log }: { items: { label: string; ms?: number; note?: string }[]; log: boolean }) {
  const vals = items.filter((i) => i.ms).map((i) => i.ms!);
  if (vals.length < 2) return null;
  const max = Math.max(...vals);
  const scale = (ms: number) => (log ? Math.log10(ms + 1) / Math.log10(max + 1) : ms / max);
  return (
    <div className="bars">
      {items.filter((i) => i.ms).map((i) => (
        <div key={i.label} className="bar-row">
          <span className="bar-label">{i.label}</span>
          <span className={`bar v-${i.label.split(' ')[0].toLowerCase()}`} style={{ width: `${Math.max(2, scale(i.ms!) * 100)}%` }} />
          <span className="bar-value">{fmtMs(i.ms)}{i.note ? ` ${i.note}` : ''}</span>
        </div>
      ))}
    </div>
  );
}

function ScenarioCard({ s, history, live, label, log, available, onRun }: {
  s: Scenario; history: RunRow[]; live: Live; label: string; log: boolean; available: Record<string, boolean>;
  onRun: (s: Scenario, variants: Variant[], volume: number, extra: Record<string, unknown>) => void;
}) {
  const [volume, setVolume] = useState(s.defaultVolume);
  const [opts, setOpts] = useState<Record<string, unknown>>(Object.fromEntries((s.options ?? []).map((o) => [o.key, o.default])));
  const variants = VARIANTS.filter((v) => s.variants[v]);
  const last = (v: Variant) => history.find((r) => r.scenario === s.id && r.variant === v && isDone(r) && (!label || r.label === label)) ?? history.find((r) => r.scenario === s.id && r.variant === v && isDone(r));
  const shown = (v: Variant) => live[`${s.id}:${v}`]?.row ?? last(v);
  const fast = shown('FAST');
  const slowBase = fast?.volume ? baseline(history, s.id, 'SLOW', fast.volume) : undefined;
  const midBase = fast?.volume && s.variants.MID ? baseline(history, s.id, 'MID', fast.volume) : undefined;
  const cmp = compare(slowBase, fast && isDone(fast) ? fast : undefined);
  const cmpMid = compare(midBase, fast && isDone(fast) ? fast : undefined);
  const anyRunning = variants.some((v) => live[`${s.id}:${v}`] && !live[`${s.id}:${v}`].finished);
  const missing = variants.filter((v) => available[s.variants[v]!.flow] === false);

  if (s.driver) return <Uc7Card s={s} volume={volume} setVolume={setVolume} label={label} live={live} available={available} onRun={onRun} />;

  return (
    <section className="card scenario">
      <header className="scenario-head">
        <div>
          <h2><span className="uc">{s.id}</span> {s.title}</h2>
          <p className="muted small">{s.theme} · KPI: {s.kpi}</p>
        </div>
        <div className="controls">
          <label>Volume
            <select value={volume} onChange={(e) => setVolume(Number(e.target.value))}>
              {[...new Set([...s.volumes, volume])].sort((a, b) => a - b).map((v) => <option key={v} value={v}>{fmtNum(v)}</option>)}
            </select>
          </label>
          {(s.options ?? []).map((o) => o.kind === 'boolean' ? (
            <label key={o.key} className="check"><input type="checkbox" checked={!!opts[o.key]} onChange={(e) => setOpts({ ...opts, [o.key]: e.target.checked })} /> {o.label}</label>
          ) : (
            <label key={o.key}>{o.label}
              <select value={Number(opts[o.key])} onChange={(e) => setOpts({ ...opts, [o.key]: Number(e.target.value) })}>
                {(o.choices ?? []).map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </label>
          ))}
          <div className="buttons">
            {variants.map((v) => <button key={v} className={`b-${v.toLowerCase()}`} disabled={anyRunning} onClick={() => onRun(s, [v], volume, opts)}>Run {v}</button>)}
            <button className="race" disabled={anyRunning} onClick={() => onRun(s, variants, volume, opts)}>Race all</button>
          </div>
        </div>
      </header>
      {s.note && <p className="note small">{s.note}</p>}
      {missing.length > 0 && <p className="bad small">Not in this build: {missing.map((v) => s.variants[v]!.flow).join(', ')} (pa app add flow).</p>}
      <div className="tiles">
        {variants.map((v) => <Tile key={v} variant={v} technique={s.variants[v]!.technique} live={live[`${s.id}:${v}`]} last={last(v)} />)}
      </div>
      {cmp.speedup != null && (
        <div className="verdict">
          <span className="speedup">⚡ {cmp.speedup >= 10 ? cmp.speedup.toFixed(0) : cmp.speedup.toFixed(1)}× faster{cmp.extrapolated ? ' (SLOW extrapolated linearly)' : ''}</span>
          {cmpMid.speedup != null && <span>vs MID {cmpMid.speedup.toFixed(1)}×</span>}
          {cmp.checksumMatch != null && <span className={cmp.checksumMatch ? 'ok' : 'bad'}>{cmp.checksumMatch ? '✅ same result' : '❌ results differ'}</span>}
          {cmp.actionsSaved != null && <span>{fmtNum(cmp.actionsSaved)} actions saved</span>}
          {cmp.slowPpr != null && cmp.fastPpr != null && <span className="muted">daily PPR: {cmp.slowPpr.toFixed(2)} % → {cmp.fastPpr.toFixed(3)} %</span>}
        </div>
      )}
      <Bars log={log} items={[
        ...(slowBase ? [{ label: `SLOW @${fmtNum(fast?.volume)}`, ms: slowBase.durationMs, note: slowBase.extrapolated ? `(extrapolated from ${fmtNum(slowBase.row.volume)})` : undefined }] : []),
        ...(midBase ? [{ label: `MID @${fmtNum(fast?.volume)}`, ms: midBase.durationMs, note: midBase.extrapolated ? `(extrapolated from ${fmtNum(midBase.row.volume)})` : undefined }] : []),
        ...(fast && isDone(fast) ? [{ label: `FAST @${fmtNum(fast.volume)}`, ms: fast.durationMs }] : []),
      ]} />
    </section>
  );
}

function Uc7Card({ s, volume, setVolume, label, live, available, onRun }: {
  s: Scenario; volume: number; setVolume: (v: number) => void; label: string; live: Live; available: Record<string, boolean>;
  onRun: (s: Scenario, variants: Variant[], volume: number, extra: Record<string, unknown>) => void;
}) {
  const driver = live[`${s.id}:DRIVER`];
  const [counts, setCounts] = useState<{ SLOW: number; FAST: number; relevantSlow: number }>();
  useEffect(() => {
    if (!driver?.correlationId || !driver.row) return;
    let stop = false;
    const tick = async () => {
      try { const c = await getTriggeredRuns(driver.correlationId); if (!stop) setCounts(c); } catch { /* keep last */ }
    };
    void tick();
    const t = setInterval(tick, 10_000);
    const end = setTimeout(() => clearInterval(t), 8 * 60_000);
    return () => { stop = true; clearInterval(t); clearTimeout(end); };
  }, [driver?.correlationId, driver?.row]);
  const expected = Math.ceil(volume / 10);
  return (
    <section className="card scenario">
      <header className="scenario-head">
        <div>
          <h2><span className="uc">{s.id}</span> {s.title}</h2>
          <p className="muted small">{s.theme} · KPI: {s.kpi}</p>
        </div>
        <div className="controls">
          <label>Modifications
            <select value={volume} onChange={(e) => setVolume(Number(e.target.value))}>{s.volumes.map((v) => <option key={v} value={v}>{v}</option>)}</select>
          </label>
          <div className="buttons">
            <button className="race" disabled={(driver && !driver.finished) || available[s.driver!] === false} onClick={() => onRun(s, [], volume, { label })}>Run driver</button>
          </div>
        </div>
      </header>
      {s.note && <p className="note small">{s.note}</p>}
      <div className="tiles">
        {(['SLOW', 'FAST'] as Variant[]).map((v) => (
          <div key={v} className={`tile v-${v.toLowerCase()}`}>
            <div className="tile-head"><strong>{v}</strong></div>
            <p className="technique">{s.variants[v]!.technique}</p>
            <dl className="kpis">
              <dt>Runs triggered</dt><dd className="big">{counts ? counts[v as 'SLOW' | 'FAST'] : '–'}</dd>
              <dt>Relevant</dt><dd>{counts ? (v === 'SLOW' ? counts.relevantSlow : counts.FAST) : '–'} <span className="muted">of {expected} expected</span></dd>
            </dl>
          </div>
        ))}
        <div className="tile">
          <div className="tile-head"><strong>DRIVER</strong>{driver?.row?.status && <span className={`status s-${driver.row.status.toLowerCase()}`}>{driver.row.status}</span>}</div>
          <p className="technique">Changes {volume} items of UC7Items in one $batch (1 in 10 → Approved).</p>
          <dl className="kpis"><dt>Duration</dt><dd>{fmtMs(driver?.row?.durationMs)}</dd><dt>Modified | relevant</dt><dd className="mono">{driver?.row?.checksum ?? '–'}</dd></dl>
          <RunLinks row={driver?.row} />
        </div>
      </div>
      {counts && counts.SLOW > 0 && <div className="verdict"><span className="speedup">{counts.SLOW} runs → {counts.FAST} runs</span><span className="muted">{counts.SLOW - counts.FAST} runs (and their billed actions) avoided by trigger conditions</span></div>}
    </section>
  );
}

function Scoreboard({ history, refresh, loading }: { history: RunRow[]; refresh: () => void; loading: boolean }) {
  const [scenario, setScenario] = useState('');
  const [labelFilter, setLabelFilter] = useState('');
  const rows = history.filter((r) => r.scenario !== 'POC' && !(r.scenario === 'UC7' && r.variant !== 'DRIVER'))
    .filter((r) => !scenario || r.scenario === scenario)
    .filter((r) => !labelFilter || (r.label ?? '').toLowerCase().includes(labelFilter.toLowerCase()));
  return (
    <section className="card">
      <header className="scenario-head">
        <h2>Scoreboard</h2>
        <div className="controls">
          <label>Scenario<select value={scenario} onChange={(e) => setScenario(e.target.value)}><option value="">all</option>{[...SCENARIOS.map((s) => s.id), 'ADMIN'].map((s) => <option key={s}>{s}</option>)}</select></label>
          <label>Label<input value={labelFilter} onChange={(e) => setLabelFilter(e.target.value)} placeholder="filter" /></label>
          <button onClick={refresh} disabled={loading}>{loading ? 'Loading…' : 'Refresh'}</button>
        </div>
      </header>
      <div className="table-wrap">
        <table className="score">
          <thead><tr><th>Started</th><th>UC</th><th>Variant</th><th>Volume</th><th>Status</th><th>Duration</th><th>Actions</th><th>Calls</th><th>Records</th><th>Payload</th><th>Checksum</th><th>Label</th><th /></tr></thead>
          <tbody>
            {rows.slice(0, 300).map((r) => (
              <tr key={r.correlationId}>
                <td className="nowrap">{r.startTime ? new Date(r.startTime).toLocaleString() : '–'}</td>
                <td>{r.scenario}</td><td>{r.variant}</td><td className="num">{fmtNum(r.volume)}</td>
                <td><span className={`status s-${(r.status ?? '').toLowerCase()}`}>{r.status}</span></td>
                <td className="num">{fmtMs(r.durationMs)}</td><td className="num">{fmtNum(r.actionsEstimated)}</td><td className="num">{fmtNum(r.connectorCalls)}</td>
                <td className="num">{fmtNum(r.recordsProcessed)}</td><td className="num">{fmtKb(r.payloadBytes)}</td>
                <td className="mono small">{r.checksum}</td><td className="small">{r.label}</td>
                <td className="nowrap"><RunLinks row={r} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

export function AdminPanel({ available }: { available: Record<string, boolean> }) {
  const [runs, setRuns] = useState<Record<string, TrackedRun>>({});
  const [target, setTarget] = useState('all');
  const run = (flow: string, volume: number) => {
    void launchAndTrack(flow, { volume, label: 'admin', ...(flow === 'RESET' ? { target } : {}) }, (r) => setRuns((x) => ({ ...x, [flow]: r })), { timeoutMs: LONG, pollMs: 5000 })
      .catch((e: unknown) => alert(e instanceof Error ? e.message : String(e)));
  };
  return (
    <section className="card">
      <h2>Seed &amp; reset</h2>
      <div className="admin">
        {ADMIN_FLOWS.map((a) => {
          const r = runs[a.flow];
          return (
            <div key={a.flow} className="tile">
              <div className="tile-head"><strong>{a.title}</strong>{r?.row?.status && <span className={`status s-${r.row.status.toLowerCase()}`}>{r.row.status}</span>}</div>
              <p className="technique">{a.description}</p>
              {a.flow === 'RESET' && (
                <label className="small">Target <select value={target} onChange={(e) => setTarget(e.target.value)}>{['all', 'uc4', 'uc5', 'uc7', 'uc9'].map((t) => <option key={t}>{t}</option>)}</select></label>
              )}
              <button disabled={(r && !r.finished) || available[a.flow] === false} onClick={() => run(a.flow, a.volume)}>{r && !r.finished ? `Running… ${fmtMs(Date.now() - r.startedAt)}` : `Run ${a.flow}`}</button>
              {r?.row?.checksum && <p className="small mono">{r.row.checksum}</p>}
              {r?.row?.error && <p className="bad small">{r.row.error.slice(0, 300)}</p>}
              <RunLinks row={r?.row} />
            </div>
          );
        })}
      </div>
    </section>
  );
}

export default function LabPage({ label }: { label: string }) {
  const [history, setHistory] = useState<RunRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [live, setLive] = useState<Live>({});
  const [log, setLog] = useState(true);
  const [available, setAvailable] = useState<Record<string, boolean>>({});
  const [, setTick] = useState(0);

  const refresh = useCallback(async () => {
    setLoading(true);
    try { setHistory(await getHistory({ top: 1000 })); } catch (e) { console.warn(e); } finally { setLoading(false); }
  }, []);

  useEffect(() => {
    void refresh();
    const flows = [...SCENARIOS.flatMap((s) => [...Object.values(s.variants).map((v) => v!.flow), ...(s.driver ? [s.driver] : [])]), ...ADMIN_FLOWS.map((a) => a.flow)];
    void Promise.all(flows.map(async (f) => [f, await hasFlow(f)] as const)).then((xs) => setAvailable(Object.fromEntries(xs)));
  }, [refresh]);

  // re-render running timers once a second
  const running = Object.values(live).some((r) => !r.finished);
  useEffect(() => {
    if (!running) return;
    const t = setInterval(() => setTick((x) => x + 1), 1000);
    return () => clearInterval(t);
  }, [running]);

  const onRun = useCallback((s: Scenario, variants: Variant[], volume: number, extra: Record<string, unknown>) => {
    const targets = s.driver ? [{ key: 'DRIVER', flow: s.driver }] : variants.map((v) => ({ key: v, flow: s.variants[v]!.flow }));
    for (const t of targets) {
      void launchAndTrack(t.flow, { volume, label, ...extra }, (r) => {
        setLive((x) => ({ ...x, [`${s.id}:${t.key}`]: { ...r, scenario: s.id, variant: t.key, volume } }));
        if (r.finished) void refresh();
      }, { timeoutMs: LONG, pollMs: 2000 }).catch((e: unknown) => {
        setLive((x) => ({ ...x, [`${s.id}:${t.key}`]: { correlationId: '-', flow: t.flow, payload: {}, startedAt: Date.now(), invoke: { invokeError: e instanceof Error ? e.message : String(e) }, finished: true, scenario: s.id, variant: t.key, volume } }));
      });
    }
  }, [label, refresh]);

  const groups = useMemo(() => [...new Set(SCENARIOS.map((s) => s.theme))], []);
  return (
    <>
      <div className="toolbar">
        <label className="check"><input type="checkbox" checked={log} onChange={(e) => setLog(e.target.checked)} /> log scale bars</label>
        <span className="muted small">{history.length} runs in history</span>
      </div>
      {groups.map((g) => (
        <div key={g} className="group">
          <h3 className="group-title">{g}</h3>
          {SCENARIOS.filter((s) => s.theme === g).map((s) => (
            <ScenarioCard key={s.id} s={s} history={history} live={live} label={label} log={log} available={available} onRun={onRun} />
          ))}
        </div>
      ))}
      <Scoreboard history={history} refresh={refresh} loading={loading} />
    </>
  );
}
