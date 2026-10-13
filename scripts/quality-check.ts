#!/usr/bin/env bun
/**
 * Quality gate: nothing reaches the client's Cloud Run unless it passes this.
 *
 *   bun run quality:check                    full gate (about 40 min: four apps run in parallel)
 *   bun run quality:check --update-baseline  accept the current results as the new baseline
 *   bun run quality:check --skip-live        static checks only (never valid for a deploy)
 *
 * 1. Static: typecheck, lint, unit tests (a failing test gets one retry: known timing flake),
 *    brand guard.
 * 2. Live: every scenario in scenarios/*.json against real Gemini Enterprise, one process per app.
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
function capabilityOf(r: ScenarioResult, fileOf: Map<string, string>): string {
  const m = /^mx-[a-z]+-([a-z]+)-\d+$/.exec(r.id);
  return m ? `${r.surface}:${m[1]}` : `${r.surface}:${fileOf.get(r.id) ?? 'other'}`;
}

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s.length ? s[Math.floor(s.length / 2)]! : 0;
}

async function runLive(runDir: string): Promise<void> {
  mkdirSync(runDir, { recursive: true });
  await Promise.all(
    SURFACES.map(
      (surface) =>
        new Promise<void>((resolve) => {
          console.log(`live: ${surface} …`);
          const child = spawn('caffeinate', ['-i', 'bun', 'run', 'scenarios'], {
            cwd: ROOT,
            env: {
              ...process.env,
              GE_SCENARIO_SURFACE: surface,
              GE_SCENARIO_RUN_DIR: join(runDir, surface),
            },
            stdio: 'ignore',
          });
          child.on('exit', () => {
            console.log(`live: ${surface} done`);
            resolve();
          });
        }),
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

  const runDir = join(RESULTS, `quality-${new Date().toISOString().replace(/[:.]/g, '-')}`);
  await runLive(runDir);

  const fileOf = new Map<string, string>();
  for (const f of readdirSync(join(ROOT, 'scenarios')).filter((n) => n.endsWith('.json'))) {
    if (f === 'quality-baseline.json') continue;
    const list = JSON.parse(readFileSync(join(ROOT, 'scenarios', f), 'utf8')) as Array<{
      id: string;
    }>;
    for (const s of list) fileOf.set(s.id, f.replace(/\.json$/, ''));
  }
  const results: ScenarioResult[] = SURFACES.flatMap((surface) => {
    const dir = join(runDir, surface);
    if (!existsSync(dir)) return [];
    return readdirSync(dir)
      .filter((n) => n.endsWith('.json') && n !== 'summary.json')
      .map((n) => JSON.parse(readFileSync(join(dir, n), 'utf8')) as ScenarioResult);
  });
  const expected = fileOf.size;

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
  const regressions = table.filter(
    (row) => baseline[row.capability] !== undefined && row.passed < baseline[row.capability]! - TOLERANCE,
  );
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
      (missing ? `, ${missing} scenario(s) produced no result` : ''),
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
