// Pairing and speed-up maths for the scoreboard (docs/03-code-app-benchmark-harness.md, "The app computes").
import type { RunRow } from '../lib/benchmark';
import type { Variant } from './scenarios';
import { PPR_PER_DAY } from './scenarios';

export const isDone = (r?: RunRow) => !!r?.status && r.status !== 'Running' && r.status !== 'PollError';

export interface Baseline {
  row: RunRow;
  /** Duration scaled linearly to the target volume when the baseline ran at another volume. */
  durationMs: number;
  actions?: number;
  extrapolated: boolean;
}

/** Latest finished run of scenario/variant at `volume`; else the nearest volume, scaled linearly. */
export function baseline(rows: RunRow[], scenario: string, variant: Variant, volume: number): Baseline | undefined {
  const done = rows.filter((r) => r.scenario === scenario && r.variant === variant && isDone(r) && r.durationMs != null && r.volume);
  const same = done.find((r) => r.volume === volume);
  if (same) return { row: same, durationMs: same.durationMs!, actions: same.actionsEstimated, extrapolated: false };
  const nearest = [...done].sort((a, b) => Math.abs(Math.log((a.volume ?? 1) / volume)) - Math.abs(Math.log((b.volume ?? 1) / volume)))[0];
  if (!nearest) return undefined;
  const f = volume / (nearest.volume ?? volume);
  return { row: nearest, durationMs: nearest.durationMs! * f, actions: nearest.actionsEstimated == null ? undefined : Math.round(nearest.actionsEstimated * f), extrapolated: true };
}

export interface Comparison {
  speedup?: number;
  extrapolated: boolean;
  checksumMatch?: boolean;
  actionsSaved?: number;
  slowPpr?: number;
  fastPpr?: number;
}

export function compare(slow: Baseline | undefined, fast: RunRow | undefined): Comparison {
  if (!slow || !fast?.durationMs) return { extrapolated: false };
  return {
    speedup: slow.durationMs / Math.max(fast.durationMs, 1),
    extrapolated: slow.extrapolated,
    checksumMatch: slow.extrapolated || !slow.row.checksum || !fast.checksum ? undefined : slow.row.checksum === fast.checksum,
    actionsSaved: slow.actions != null && fast.actionsEstimated != null ? slow.actions - fast.actionsEstimated : undefined,
    slowPpr: slow.actions == null ? undefined : (slow.actions / PPR_PER_DAY) * 100,
    fastPpr: fast.actionsEstimated == null ? undefined : (fast.actionsEstimated / PPR_PER_DAY) * 100,
  };
}

export const fmtMs = (ms?: number) => {
  if (ms == null) return '–';
  if (ms < 1000) return `${Math.round(ms)} ms`;
  if (ms < 120_000) return `${(ms / 1000).toFixed(1)} s`;
  if (ms < 7_200_000) return `${(ms / 60_000).toFixed(1)} min`;
  return `${(ms / 3_600_000).toFixed(1)} h`;
};
export const fmtNum = (n?: number) => (n == null ? '–' : n.toLocaleString('en-US'));
export const fmtKb = (b?: number) => (b == null || b === 0 ? '–' : b < 1024 * 1024 ? `${(b / 1024).toFixed(0)} KB` : `${(b / 1024 / 1024).toFixed(1)} MB`);
