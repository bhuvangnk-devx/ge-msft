// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  defaultExcelSeed,
  installFakeExcel,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type MountedStack,
} from '../test-harness/index.js';
import { classifyFailure } from '../reporter.js';

/**
 * Live 2026-10-14: a write to a protected sheet ran 12 turns and ended as "That didn't finish" with
 * Try again, never saying the sheet was protected. A request the plugin cannot do (pivot table)
 * ended the same way. Both now stop at once and say what to do, with no error card.
 */

let sim: ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

async function runTask(task: string, turns: string[]) {
  const client = scriptedClient(turns);
  ui = mountStack({ surface: 'excel', client });
  await ui.flush();
  const since = new Date().toISOString();
  let run!: Promise<void>;
  await ui.act(() => {
    run = ui!.controller.runCommands(task);
  });
  await ui.waitFor((s) => !s.busy || s.pendingPlan !== undefined);
  if (ui.controller.getState().pendingPlan) await ui.act(() => ui!.controller.approvePlan());
  await ui.waitFor((s) => !s.busy);
  await run;
  await ui.flush();
  return { since, reply: ui.controller.getState().messages.at(-1)! };
}

const write = '```cmd\nset Summary!A1 Region\n```';

describe('a task the user must unblock', () => {
  it('a protected sheet: one turn, says it is protected, no error card, no report', async () => {
    const seed = defaultExcelSeed();
    seed.sheets.find((s) => s.name === 'Summary')!.protected = true;
    sim = installFakeExcel(seed);
    const { since, reply } = await runTask('Put Region in A1 of Summary', [write, write, write]);

    expect(reply.error).toBeUndefined();
    expect(reply.text).toContain('Summary!A1 is protected or read-only');
    expect(reply.text).toContain('Unprotect Sheet');
    expect(ui!.container.querySelector('.error-card')).toBeNull();
    expect(ui!.controller.getState().steps.filter((s) => s.kind === 'turn-start')).toHaveLength(1);
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[0]?.[0]).not.toBe(
      'Region',
    );
    expect(
      classifyFailure(JSON.parse(ui!.controller.diagnostics({ surface: 'excel' })), since),
    ).toBeUndefined();
  });

  it('an unprotected sheet still takes the write (no regression)', async () => {
    sim = installFakeExcel();
    const { reply } = await runTask('Put Region in A1 of Summary', [write, '```cmd\ndone\n```']);
    expect(reply.error).toBeUndefined();
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[0]?.[0]).toBe('Region');
  });

  it('a pivot table: says it is not supported yet and what it can do instead', async () => {
    sim = installFakeExcel();
    const pivot =
      '```cmd\n/insert-pivot sourceRange=Sales!A1:D7 destinationRange=Summary!E1 rowFields=region valueFields=revenue\n```';
    const { since, reply } = await runTask('Make a pivot table of revenue by region', [
      pivot,
      pivot,
      pivot,
    ]);

    expect(reply.error).toBeUndefined();
    expect(reply.text).toContain("CN-GPT can't do this yet: Pivot table.");
    expect(reply.text).toContain('I can help with: cells');
    expect(ui!.container.querySelector('.error-card')).toBeNull();
    expect(ui!.controller.getState().steps.filter((s) => s.kind === 'turn-start')).toHaveLength(1);
    expect(
      classifyFailure(JSON.parse(ui!.controller.diagnostics({ surface: 'excel' })), since),
    ).toBeUndefined();
  });

  it('answers in Bahasa when the user writes in Bahasa', async () => {
    sim = installFakeExcel();
    const pivot =
      '```cmd\n/insert-pivot sourceRange=Sales!A1:D7 destinationRange=Summary!E1 rowFields=region valueFields=revenue\n```';
    const { reply } = await runTask('Tolong buatkan pivot table untuk data ini', [pivot]);
    expect(reply.text).toContain('CN-GPT belum bisa melakukan ini: Pivot table.');
  });
});
