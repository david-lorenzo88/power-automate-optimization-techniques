// Launch a benchmark/PoC flow and follow its telemetry row in pa_benchmarkrun.
import { findFlow, findTable } from './services';

export const PREFIX: string = import.meta.env.VITE_PREFIX ?? 'pa';
const col = (name: string) => `${PREFIX}_${name}`;

export interface RunRow {
  correlationId: string;
  scenario?: string;
  variant?: string;
  status?: string;
  durationMs?: number;
  volume?: number;
  recordsProcessed?: number;
  connectorCalls?: number;
  actionsEstimated?: number;
  payloadBytes?: number;
  runId?: string;
  flowId?: string;
  environmentId?: string;
  runUrl?: string;
  runUrlSolution?: string;
  editUrl?: string;
  details?: string;
  error?: string;
}

export function toRunRow(r: Record<string, unknown>): RunRow {
  const s = (k: string) => (r[col(k)] ?? undefined) as string | undefined;
  const n = (k: string) => (r[col(k)] == null ? undefined : Number(r[col(k)]));
  return {
    correlationId: s('name') ?? '',
    scenario: s('scenario'),
    variant: s('variant'),
    status: s('status'),
    durationMs: n('durationms'),
    volume: n('volume'),
    recordsProcessed: n('recordsprocessed'),
    connectorCalls: n('connectorcalls'),
    actionsEstimated: n('actionsestimated'),
    payloadBytes: n('payloadbytes'),
    runId: s('runid'),
    flowId: s('flowid'),
    environmentId: s('environmentid'),
    runUrl: s('runurl'),
    runUrlSolution: s('runurlsolution'),
    editUrl: s('editurl'),
    details: s('details'),
    error: s('error'),
  };
}

export async function getRunRow(correlationId: string): Promise<RunRow | undefined> {
  const table = await findTable('benchmarkrun');
  if (!table) throw new Error('pa_benchmarkrun service not generated. Run: pa app add data-source --connector dataverse --table pa_benchmarkrun');
  const res = await table.getAll({ filter: `${col('name')} eq '${correlationId}'`, top: 1 });
  if (!res.success) throw res.error ?? new Error('getAll failed');
  const row = res.data?.[0];
  return row ? toRunRow(row) : undefined;
}

export interface InvokeOutcome {
  invokeMs?: number;
  invokeSuccess?: boolean;
  invokeError?: string;
  responseData?: unknown;
}

export interface TrackedRun {
  correlationId: string;
  flow: string;
  payload: Record<string, unknown>;
  startedAt: number;
  invoke: InvokeOutcome;
  row?: RunRow;
  finished: boolean;
  timedOut?: boolean;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/**
 * Starts the flow WITHOUT blocking on its response (async pattern) and polls the
 * telemetry row until it leaves "Running" and the invoke promise has settled.
 */
export async function launchAndTrack(
  flowNamePart: string,
  payload: Record<string, unknown>,
  onUpdate: (r: TrackedRun) => void,
  { pollMs = 2000, timeoutMs = 6 * 60_000 } = {},
): Promise<TrackedRun> {
  const run = await findFlow(flowNamePart);
  if (!run) throw new Error(`Flow service for "${flowNamePart}" not generated. Run: pa app add flow --flow-id <id>`);

  const correlationId = crypto.randomUUID();
  const fullPayload = { correlationId, ...payload };
  const payloadText = JSON.stringify(fullPayload);
  const tracked: TrackedRun = { correlationId, flow: flowNamePart, payload: fullPayload, startedAt: Date.now(), invoke: {}, finished: false };
  const emit = () => onUpdate({ ...tracked, invoke: { ...tracked.invoke } });
  emit();

  let invokeSettled = false;
  // Single text input on the trigger. We pass it under both the raw key ("text")
  // and the title ("payload"), because generators differ in which one they expose.
  run({ text: payloadText, payload: payloadText })
    .then((res) => {
      tracked.invoke = {
        invokeMs: Date.now() - tracked.startedAt,
        invokeSuccess: res.success,
        invokeError: res.success ? undefined : String(res.error?.message ?? res.error),
        responseData: res.data,
      };
    })
    .catch((e: unknown) => {
      tracked.invoke = { invokeMs: Date.now() - tracked.startedAt, invokeSuccess: false, invokeError: e instanceof Error ? e.message : String(e) };
    })
    .finally(() => {
      invokeSettled = true;
      emit();
    });

  while (Date.now() - tracked.startedAt < timeoutMs) {
    await sleep(pollMs);
    try {
      tracked.row = await getRunRow(correlationId);
    } catch (e) {
      tracked.row = { correlationId, status: 'PollError', error: e instanceof Error ? e.message : String(e) };
    }
    emit();
    const done = tracked.row?.status && tracked.row.status !== 'Running' && tracked.row.status !== 'PollError';
    if (done && invokeSettled) break;
  }
  tracked.finished = true;
  tracked.timedOut = !(tracked.row?.status && tracked.row.status !== 'Running');
  emit();
  return tracked;
}
