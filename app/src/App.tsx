import { useEffect, useState } from 'react';
import { getContext } from '@microsoft/power-apps/app';
import type { IContext } from '@microsoft/power-apps/app';
import LabPage, { AdminPanel } from './lab/LabPage';
import { SCENARIOS, ADMIN_FLOWS } from './lab/scenarios';
import { hasFlow } from './lib/services';
import PocPage from './PocPage';
import './App.css';

type Tab = 'lab' | 'admin' | 'poc';

const readTab = (): Tab => {
  try { const t = localStorage.getItem('fol.tab'); return t === 'admin' || t === 'poc' ? t : 'lab'; } catch { return 'lab'; }
};

export default function App() {
  const [tab, setTab] = useState<Tab>(readTab);
  const [ctx, setCtx] = useState<IContext>();
  const [label, setLabel] = useState(() => { try { return localStorage.getItem('fol.label') ?? 'live'; } catch { return 'live'; } });
  const [available, setAvailable] = useState<Record<string, boolean>>({});

  useEffect(() => {
    getContext().then(setCtx).catch(() => undefined);
    const flows = [...SCENARIOS.flatMap((s) => [...Object.values(s.variants).map((v) => v!.flow), ...(s.driver ? [s.driver] : [])]), ...ADMIN_FLOWS.map((a) => a.flow)];
    void Promise.all(flows.map(async (f) => [f, await hasFlow(f)] as const)).then((xs) => setAvailable(Object.fromEntries(xs)));
  }, []);
  useEffect(() => { try { localStorage.setItem('fol.tab', tab); } catch { /* private mode */ } }, [tab]);
  useEffect(() => { try { localStorage.setItem('fol.label', label); } catch { /* private mode */ } }, [label]);

  return (
    <main className="lab">
      <header className="app-head">
        <div>
          <h1>Flow Optimization Lab</h1>
          <p className="muted small">Same input, same result, different cost: slow vs fast Power Automate patterns{ctx ? ` · ${ctx.user.fullName}` : ''}</p>
        </div>
        <label className="label-input">Run label<input value={label} onChange={(e) => setLabel(e.target.value)} placeholder="e.g. Dubai live demo" /></label>
      </header>
      <nav className="tabs">
        {(['lab', 'admin', 'poc'] as Tab[]).map((t) => (
          <button key={t} className={tab === t ? 'on' : ''} onClick={() => setTab(t)}>{t === 'lab' ? 'Scenarios' : t === 'admin' ? 'Seed & reset' : 'PoC checks'}</button>
        ))}
      </nav>
      {tab === 'lab' && <LabPage label={label} />}
      {tab === 'admin' && <AdminPanel available={available} />}
      {tab === 'poc' && <PocPage />}
    </main>
  );
}
