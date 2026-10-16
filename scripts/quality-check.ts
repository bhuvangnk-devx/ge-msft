#!/usr/bin/env bun
/**
 * Quality gate: nothing reaches the client's Cloud Run unless it passes this.
 *
 *   bun run quality:check                    daily sample: QUALITY_SAMPLE=4 scenarios per capability
 *   bun run quality:check --full             every scenario (about a whole day's Gemini pool)
 *   bun run quality:check --update-baseline  accept the current results as the new baseline
 *   bun run quality:check --skip-live        static checks only (never valid for a deploy)
 *
 * 1. Static: typecheck, lint, unit tests (a failing test gets one retry: known timing flake),
 *    brand guard.
 * 2. Live: scenarios/*.json against real Gemini Enterprise, QUALITY_SHARDS processes per app.
 *    Gemini Enterprise Standard allows 160 assistant queries a day per licence, pooled with CIMB's
 *    users (about 2,400 a day; resets at midnight PT = 14:00 WIB). A scenario costs about 3.2, so the
 *    full set (~700) spends the day's pool. The daily sample (~175 scenarios, ~570 queries) rotates
 *    which scenarios run, so every one is covered every few days. The first quota 429 stops the
 *    run; the rest are reported as not run, never as failures.
 * 3. Compare each capability with scenarios/quality-baseline.json. The gate fails when a
 *    capability passes fewer scenarios than its baseline allows, or the overall rate is under 95%.
 *
 * Writes scenarios/.results/quality-check.json (commit, verdict, per-capability table). The
 * deploy script refuses a commit without a passing report for that exact commit.
 */
import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const ROOT = join(import.meta.dir, '..');
const RESULTS = join(ROOT, 'scenarios', '.results');
const BASELINE = join(ROOT, 'scenarios', 'quality-baseline.json');
const REPORT = join(RESULTS, 'quality-check.json');
const SURFACES = ['excel', 'word', 'powerpoint', 'outlook'] as const;
/** A capability may lose this many passes to model variance before the gate fails. */
const TOLERANCE = 1;
const MIN_OVERALL = 0.95;
/**
 * Parallel processes per app (QUALITY_SHARDS, default 1). The project allows 30,000 agent-search
 * tokens a minute, shared with CIMB staging: more than 4 processes in total got 429s (2026-10-15).
 */
const SHARDS = Math.min(8, Math.max(1, Math.trunc(Number(process.env.QUALITY_SHARDS ?? 1)) || 1));

const args = new Set(process.argv.slice(2));
const sh = (cmd: string, cmdArgs: string[]): boolean => {
  console.log(`\n$ ${cmd} ${cmdArgs.join(' ')}`);
  return spawnSync(cmd, cmdArgs, { cwd: ROOT, stdio: 'inherit' }).status === 0;
};

interface ScenarioResult {
  id: string;
  surface: string;
  passed: boolean;
  durationMs: number;
  failures: string[];
  transcript?: unknown[];
}

/** Group key: matrix ids are mx-<app>-<capability>-NN; any other scenario counts per file. */
function capabilityOf(r: { id: string; surface: string }, fileOf: Map<string, string>): string {
  const m = /^mx-[a-z]+-([a-z]+)-\d+$/.exec(r.id);
  return m ? `${r.surface}:${m[1]}` : `${r.surface}:${fileOf.get(r.id) ?? 'other'}`;
}

/** Scenarios per capability in the daily sample (QUALITY_SAMPLE); `--full` runs them all. */
const SAMPLE = Math.max(1, Math.trunc(Number(process.env.QUALITY_SAMPLE ?? 4)) || 4);

/**
 * The day's sample: SAMPLE consecutive scenarios per capability, starting further along each day
 * (wrapping), so consecutive days cover different scenarios.
 */
function dailySample(
  all: Array<{ id: string; surface: string }>,
  fileOf: Map<string, string>,
): string[] {
  const day = Math.floor(Date.now() / 86_400_000);
  const groups = new Map<string, string[]>();
  for (const s of all) {
    const key = capabilityOf(s, fileOf);
    groups.set(key, [...(groups.get(key) ?? []), s.id]);
  }
  return [...groups.values()].flatMap((ids) => {
    if (ids.length <= SAMPLE) return ids;
    const start = (day * SAMPLE) % ids.length;
    return Array.from({ length: SAMPLE }, (_, i) => ids[(start + i) % ids.length]!);
  });
}

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s.length ? s[Math.floor(s.length / 2)]! : 0;
}

async function runLive(runDir: string, ids: string[] | undefined): Promise<void> {
  mkdirSync(runDir, { recursive: true });
  await Promise.all(
    SURFACES.flatMap((surface) =>
      Array.from(
        { length: SHARDS },
        (_, shard) =>
          new Promise<void>((resolve) => {
            console.log(`live: ${surface} ${shard + 1}/${SHARDS} …`);
            const child = spawn('caffeinate', ['-i', 'bun', 'run', 'scenarios'], {
              cwd: ROOT,
              env: {
                ...process.env,
                GE_SCENARIO_SURFACE: surface,
                GE_SCENARIO_SHARD: `${shard}/${SHARDS}`,
                GE_SCENARIO_RUN_DIR: join(runDir, surface),
                ...(ids ? { GE_SCENARIO_IDS: ids.join(',') } : {}),
              },
              stdio: 'ignore',
            });
            child.on('exit', () => {
              console.log(`live: ${surface} ${shard + 1}/${SHARDS} done`);
              resolve();
            });
          }),
      ),
    ),
  );
}

async function main(): Promise<void> {
  const commit = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: ROOT }).stdout.toString().trim();
  // Tracked changes anywhere, or untracked files in what gets built (untracked notes and
  // spreadsheets at the repo root do not ship, so they do not count).
  const dirty = spawnSync(
    'git',
    ['status', '--porcelain', '--untracked-files=all', '--', '.', ':!*.md', ':!*.docx', ':!*.xlsx', ':!*.pdf', ':!*.png', ':!*.html'],
    { cwd: ROOT },
  )
    .stdout.toString()
    .trim();
  if (dirty) console.warn('\n⚠ uncommitted changes: this report cannot authorise a deploy.');

  const statics: Array<[string, boolean]> = [
    ['typecheck', sh('bun', ['run', 'typecheck'])],
    ['lint', sh('bun', ['run', 'lint'])],
    ['unit tests', sh('bunx', ['vitest', 'run', 'packages', '--retry=1'])],
    ['brand guard', sh('bash', ['scripts/brand-guard.sh'])],
  ];
  const staticOk = statics.every(([, ok]) => ok);
  if (!staticOk || args.has('--skip-live')) {
    for (const [name, ok] of statics) console.log(`${ok ? '✅' : '❌'} ${name}`);
    if (!staticOk) process.exit(1);
    console.log('\nStatic checks passed; live scenarios skipped (not valid for a deploy).');
    return;
  }

  const fileOf = new Map<string, string>();
  const all: Array<{ id: string; surface: string }> = [];
  for (const f of readdirSync(join(ROOT, 'scenarios')).filter((n) => n.endsWith('.json'))) {
    if (f === 'quality-baseline.json') continue;
    const list = JSON.parse(readFileSync(join(ROOT, 'scenarios', f), 'utf8')) as Array<{
      id: string;
      surface: string;
    }>;
    for (const s of list) {
      fileOf.set(s.id, f.replace(/\.json$/, ''));
      all.push(s);
    }
  }
  const full = args.has('--full');
  const sample = full ? undefined : dailySample(all, fileOf);
  console.log(
    full
      ? `\nfull run: ${all.length} scenarios (about ${Math.round(all.length * 3.2)} Gemini queries)`
      : `\ndaily sample: ${sample!.length} of ${all.length} scenarios (about ${Math.round(sample!.length * 3.2)} Gemini queries)`,
  );

  const runDir = join(RESULTS, `quality-${new Date().toISOString().replace(/[:.]/g, '-')}`);
  await runLive(runDir, sample);
  const quotaHit = existsSync(join(runDir, 'QUOTA_EXHAUSTED'));
  const results: ScenarioResult[] = SURFACES.flatMap((surface) => {
    const dir = join(runDir, surface);
    if (!existsSync(dir)) return [];
    return readdirSync(dir)
      .filter((n) => n.endsWith('.json') && n !== 'summary.json')
      .map((n) => JSON.parse(readFileSync(join(dir, n), 'utf8')) as ScenarioResult);
  });
  const expected = sample?.length ?? all.length;

  const groups = new Map<string, ScenarioResult[]>();
  for (const r of results) {
    const key = capabilityOf(r, fileOf);
    groups.set(key, [...(groups.get(key) ?? []), r]);
  }
  const table = [...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([capability, rs]) => ({
      capability,
      passed: rs.filter((r) => r.passed).length,
      total: rs.length,
      medianSeconds: Math.round(median(rs.map((r) => r.durationMs / 1000))),
      medianGeminiCalls: median(rs.map((r) => r.transcript?.length ?? 0)),
      failing: rs.filter((r) => !r.passed).map((r) => `${r.id}: ${r.failures[0] ?? ''}`),
    }));

  const baseline: Record<string, number> = existsSync(BASELINE)
    ? (JSON.parse(readFileSync(BASELINE, 'utf8')) as Record<string, number>)
    : {};
  // A sample has no per-capability baseline (the baseline counts a whole group): it allows
  // TOLERANCE failures per capability instead.
  const regressions = table.filter((row) =>
    sample
      ? row.total - row.passed > TOLERANCE
      : baseline[row.capability] !== undefined && row.passed < baseline[row.capability]! - TOLERANCE,
  );
  const queries = results.reduce((n, r) => n + (r.transcript?.length ?? 0), 0);
  const passed = results.filter((r) => r.passed).length;
  const overall = results.length ? passed / results.length : 0;
  const missing = expected - results.length;
  const ok = regressions.length === 0 && overall >= MIN_OVERALL && missing === 0;

  console.log('\nCapability                        Passed  Median  Calls');
  for (const row of table) {
    const base = baseline[row.capability];
    const flag = regressions.includes(row) ? ' ❌ below baseline ' + base : '';
    console.log(
      `${row.capability.padEnd(32)} ${`${row.passed}/${row.total}`.padStart(6)}  ${`${row.medianSeconds}s`.padStart(6)}  ${String(row.medianGeminiCalls).padStart(5)}${flag}`,
    );
  }
  console.log(
    `\nOverall ${passed}/${results.length} (${Math.round(overall * 100)}%)` +
      (missing ? `, ${missing} scenario(s) produced no result` : '') +
      ` · ${queries} Gemini queries`,
  );
  if (quotaHit)
    console.log(
      '⚠ Gemini daily query quota reached (resets 14:00 WIB): the run stopped; scenarios without a result were not run.',
    );
  console.log(ok ? '✅ QUALITY CHECK PASSED' : '❌ QUALITY CHECK FAILED');

  mkdirSync(RESULTS, { recursive: true });
  writeFileSync(
    REPORT,
    JSON.stringify(
      {
        commit,
        clean: !dirty,
        ok: ok && !dirty,
        at: new Date().toISOString(),
        mode: sample ? 'sample' : 'full',
        quotaHit,
        queries,
        overall: { passed, total: results.length, expected },
        regressions: regressions.map((r) => r.capability),
        table,
        runDir,
      },
      null,
      2,
    ),
  );
  if (args.has('--update-baseline')) {
    writeFileSync(
      BASELINE,
      JSON.stringify(Object.fromEntries(table.map((r) => [r.capability, r.passed])), null, 2) + '\n',
    );
    console.log(`baseline updated: ${BASELINE}`);
  }
  if (!ok) process.exit(1);
}

void main();
