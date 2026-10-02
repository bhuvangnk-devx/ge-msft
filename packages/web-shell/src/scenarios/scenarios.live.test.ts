// @vitest-environment jsdom
/**
 * Live feature scenarios: real Gemini Enterprise + the real pane stack over simulated Office
 * documents. Off by default (it calls the engine); run with `bun run scenarios`.
 *
 *   GE_SCENARIOS=1                       enable
 *   GE_SCENARIO=excel-format,word-reply  only these ids (substring match)
 */
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { runScenario, type Scenario, type ScenarioResult } from './runner.js';

const LIVE = process.env.GE_SCENARIOS === '1';
const REPO_ROOT = join(__dirname, '..', '..', '..', '..');
const DIR = join(REPO_ROOT, 'scenarios');
const filter = (process.env.GE_SCENARIO ?? '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

function loadScenarios(): Scenario[] {
  return readdirSync(DIR)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .flatMap((f) => JSON.parse(readFileSync(join(DIR, f), 'utf8')) as Scenario[])
    .filter((s) => !filter.length || filter.some((id) => s.id.includes(id)));
}

const scenarios = LIVE ? loadScenarios() : [];
const results: ScenarioResult[] = [];
const runDir = join(DIR, '.results', new Date().toISOString().replace(/[:.]/g, '-'));

describe.skipIf(!LIVE)('live scenarios', () => {
  for (const scenario of scenarios) {
    it(
      `${scenario.id}: ${scenario.title ?? scenario.prompt}`,
      async () => {
        const result = await runScenario(scenario);
        results.push(result);
        mkdirSync(runDir, { recursive: true });
        writeFileSync(join(runDir, `${scenario.id}.json`), JSON.stringify(result, null, 2));
        expect(result.failures, result.failures.join('\n')).toEqual([]);
      },
      (scenario.timeoutMs ?? 240_000) + 30_000,
    );
  }

  afterAll(() => {
    if (!results.length) return;
    const rows = results.map(
      (r) =>
        `| ${r.passed ? 'PASS' : 'FAIL'} | ${r.id} | ${r.status ?? '-'} | ${
          Object.entries(r.applied)
            .map(([k, n]) => `${k}×${n}`)
            .join(', ') || '-'
        } | ${(r.durationMs / 1000).toFixed(1)}s | ${r.failures.join('; ') || '-'} |`,
    );
    const summary = [
      `# Scenario run ${new Date().toISOString()}`,
      '',
      `${results.filter((r) => r.passed).length}/${results.length} passed`,
      '',
      '| Result | Scenario | Status | Applied | Time | Failures |',
      '| --- | --- | --- | --- | --- | --- |',
      ...rows,
      '',
    ].join('\n');
    mkdirSync(runDir, { recursive: true });
    writeFileSync(join(runDir, 'summary.md'), summary);
    console.log(`\n${summary}\nDetails: ${runDir}`);
  });
});
