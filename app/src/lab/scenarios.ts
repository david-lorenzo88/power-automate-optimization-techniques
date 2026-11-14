// Scenario catalogue for the lab UI (docs/02-demo-scenarios.md). Flow names must match
// lab/flows/scenarios.mjs; the app finds each flow's generated service by name.

export type Variant = 'SLOW' | 'MID' | 'FAST';

export interface ScenarioOption {
  key: string;
  label: string;
  kind: 'number' | 'boolean' | 'select';
  default: number | boolean;
  choices?: number[];
}

export interface Scenario {
  id: string;
  title: string;
  theme: string;
  kpi: string;
  variants: Partial<Record<Variant, { flow: string; technique: string }>>;
  volumes: number[];
  defaultVolume: number;
  maxVolume: number;
  options?: ScenarioOption[];
  note?: string;
  /** UC7: a driver flow; the variants are automated flows that log one row per run. */
  driver?: string;
}

export const SCENARIOS: Scenario[] = [
  {
    id: 'UC1', title: 'Data shaping', theme: 'Kill loops', kpi: 'duration, actions (≈3×N vs 3)',
    variants: {
      SLOW: { flow: 'UC1-SLOW', technique: 'Apply to each → Condition → Compose → Append to array variable' },
      FAST: { flow: 'UC1-FAST', technique: 'Filter array + Select' },
    },
    volumes: [100, 500, 1000, 5000], defaultVolume: 500, maxVolume: 5000,
  },
  {
    id: 'UC2', title: 'Join / enrichment', theme: 'Kill loops', kpi: 'duration, connector calls (N+1 vs 2)',
    variants: {
      SLOW: { flow: 'UC2-SLOW', technique: 'Apply to each → query the customer per order' },
      FAST: { flow: 'UC2-FAST', technique: 'Dictionary pattern: json(concat(join(Select))) + key lookup' },
    },
    volumes: [100, 500, 1000, 5000], defaultVolume: 200, maxVolume: 5000,
  },
  {
    id: 'UC3', title: 'Aggregation per region', theme: 'XPath', kpi: 'duration, actions, identical totals',
    variants: {
      SLOW: { flow: 'UC3-SLOW', technique: 'Nested loops + Increment / Set variable' },
      FAST: { flow: 'UC3-FAST', technique: 'xml() + xpath count/sum, min()/max(), union() group-by' },
    },
    volumes: [100, 300, 1000, 5000], defaultVolume: 100, maxVolume: 5000,
    note: 'SLOW loops over all orders once per region (5 × N iterations).',
  },
  {
    id: 'UC4', title: 'SharePoint bulk create', theme: 'Bulk', kpi: 'duration, HTTP requests (N vs ⌈N/500⌉)',
    variants: {
      SLOW: { flow: 'UC4-SLOW', technique: 'Apply to each → Create item (concurrency 1)' },
      MID: { flow: 'UC4-MID', technique: 'Same loop, concurrency 50 (watch the 429 retries)' },
      FAST: { flow: 'UC4-FAST', technique: '$batch via Send HTTP request to SharePoint (≤500 per batch)' },
    },
    volumes: [100, 500, 1000, 2000], defaultVolume: 200, maxVolume: 5000,
    options: [{ key: 'batchSize', label: 'Batch size', kind: 'select', default: 500, choices: [100, 250, 500] }],
    note: 'PoC: SharePoint stops a $batch after ~90 s of server time (≈900 items), and changesets are not atomic.',
  },
  {
    id: 'UC5', title: 'Dataverse bulk create', theme: 'Bulk', kpi: 'duration, requests, all-or-nothing',
    variants: {
      SLOW: { flow: 'UC5-SLOW', technique: 'Apply to each → Add a new row (concurrency 1)' },
      MID: { flow: 'UC5-MID', technique: 'Same loop, concurrency 50' },
      FAST: { flow: 'UC5-FAST', technique: 'CreateMultiple, 1,000 rows per request (HTTP with Entra ID)' },
    },
    volumes: [100, 500, 1000, 5000], defaultVolume: 500, maxVolume: 5000,
    options: [{ key: 'injectBadItem', label: 'Inject 1 bad row (rollback demo)', kind: 'boolean', default: false }],
    note: 'With a bad row: the loop creates N−1 rows; CreateMultiple rolls back the whole request (0 rows).',
  },
  {
    id: 'UC6', title: 'Filter at source (SharePoint)', theme: 'Filter at source', kpi: 'duration, payload KB, retrieved vs used',
    variants: {
      SLOW: { flow: 'UC6-SLOW', technique: 'Get items (all columns, paginated) → Filter array → Get item per row' },
      FAST: { flow: 'UC6-FAST', technique: 'REST $filter + $select + $expand lookup' },
    },
    volumes: [1000, 2000, 5000, 20000], defaultVolume: 2000, maxVolume: 20000,
    note: 'Volume = order window (IDs 1…N). Both variants return the EMEA + Open orders in it.',
  },
  {
    id: 'UC6DV', title: 'Aggregate at source (Dataverse)', theme: 'Filter at source', kpi: 'duration, payload KB',
    variants: {
      SLOW: { flow: 'UC6DV-SLOW', technique: 'List rows (all columns) → Filter array → sum in the flow' },
      FAST: { flow: 'UC6DV-FAST', technique: 'FetchXML aggregate (count + sum)' },
    },
    volumes: [1000, 5000, 10000, 20000], defaultVolume: 5000, maxVolume: 20000,
  },
  {
    id: 'UC7', title: 'Trigger conditions', theme: 'Filter at source', kpi: 'number of runs for N modifications',
    driver: 'UC7-DRIVER',
    variants: {
      SLOW: { flow: 'UC7-SLOW', technique: 'Runs on every change → Condition → Terminate' },
      FAST: { flow: 'UC7-FAST', technique: 'Trigger conditions: runs only when Status = Approved' },
    },
    volumes: [20, 50, 100], defaultVolume: 100, maxVolume: 100,
    note: 'The driver changes N items (1 in 10 relevant). The automated flows poll every minute: counts settle after 1–3 min.',
  },
  {
    id: 'UC8', title: 'Platform tuning traps', theme: 'Tuning', kpi: 'correctness flags',
    variants: {
      SLOW: { flow: 'UC8-SLOW', technique: 'No pagination (100 items!) + concurrent loop on a variable (race)' },
      FAST: { flow: 'UC8-FAST', technique: 'Pagination threshold + Select / xpath (no variables)' },
    },
    volumes: [500, 1000, 2000, 5000], defaultVolume: 1000, maxVolume: 5000,
    note: 'SLOW returns a WRONG checksum on purpose: that is the demo.',
  },
  {
    id: 'UC9', title: 'Excel bulk write', theme: 'Bonus', kpi: 'duration (hours → seconds)',
    variants: {
      SLOW: { flow: 'UC9-SLOW', technique: 'Apply to each → Add a row into a table' },
      FAST: { flow: 'UC9-FAST', technique: 'Office Script setValues() in one call' },
    },
    volumes: [50, 200, 500, 2000], defaultVolume: 200, maxVolume: 2000,
  },
  {
    id: 'UC10', title: 'Fan-out', theme: 'Bonus', kpi: 'wall-clock time',
    variants: {
      SLOW: { flow: 'UC10-SLOW', technique: 'One sequential loop over all records' },
      FAST: { flow: 'UC10-FAST', technique: 'chunk() → Apply to each (concurrency 10) → child flow per chunk' },
    },
    volumes: [200, 1000, 5000, 20000], defaultVolume: 1000, maxVolume: 20000,
  },
];

export const ADMIN_FLOWS = [
  { flow: 'SEED-SP', title: 'Seed SharePoint', description: 'Fills empty lists: Customers (2,000), Orders (volume), UC7Items (100).', volume: 20000 },
  { flow: 'SEED-DV', title: 'Seed Dataverse', description: 'Fills an empty pa_order table with CreateMultiple (volume rows).', volume: 20000 },
  { flow: 'RESET', title: 'Reset demo targets', description: 'uc4: empty OrderImport · uc5: BulkDelete pa_orderimport · uc7: items back to Pending · uc9: clear workbooks.', volume: 1 },
];

export const PPR_PER_DAY = 40_000; // Power Platform requests per user/day (Premium)
