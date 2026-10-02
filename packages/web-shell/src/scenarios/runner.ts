/**
 * Runs one scenario the way a user would: a seeded (simulated) Office document, the real pane stack
 * (runtime extensions → composeSession → PanelController), the real Gemini Enterprise engine, and
 * every approval gate accepted. Returns what happened plus the failed expectations.
 */
import { commandPaletteFor, intentsForManifest, type Surface } from '@ge/contracts';
import type { DocBridge, RunRecord } from '@ge/runtime';
import {
  inferImplicitIntent,
  invocationToGrounding,
  isActuating,
  shouldUsePlannerForFreeText,
} from '../taskpane/components/App.js';
import { parseComposerInput } from '../taskpane/components/Composer.js';
import { invocationToSeed } from '../taskpane/components/quick-action-seed.js';
import { extractDirectCommandProgram } from '../taskpane/direct-command.js';
import { composeSession } from '../compose.js';
import { PanelController, type PanelState } from '../controller.js';
import { connectPanelRuntime } from '../panel-runtime.js';
import { createApplicationRuntime } from '../runtime-extensions.js';
import { shellConfigFromEnv } from '../taskpane/config.js';
import { selectBridge } from '../taskpane/select-bridge.js';
import { initialUnit } from '../taskpane/unit.js';
import { excelSeed, installFakeExcel } from '../test-harness/fake-excel.js';
import { installFakeOutlook, outlookSeed } from '../test-harness/fake-outlook.js';
import { installFakePowerPoint, powerPointSeed } from '../test-harness/fake-powerpoint.js';
import { installFakeWord, wordSeed } from '../test-harness/fake-word.js';
import { liveFetch, loadShellEnv, scenarioAuth, type TranscriptEntry } from './live-env.js';

export interface ScenarioExpect {
  /** Final status of the last task (default "completed"). */
  status?: RunRecord['status'];
  /** No error step and no panel error (default true). */
  noErrors?: boolean;
  /** Minimum number of applied changes per kind, e.g. { "format-cells": 1 }. */
  applied?: Record<string, number>;
  /** Kinds that must not be applied. */
  notApplied?: string[];
  /** Case-insensitive text the final answer must contain. */
  answerIncludes?: string[];
  /** Excel: format properties at least one formatted range must carry (e.g. "fontColor"). */
  formatsInclude?: string[];
  /** Case-insensitive text that must appear somewhere in the document afterwards. */
  documentIncludes?: string[];
}

export interface Scenario {
  id: string;
  surface: Exclude<Surface, 'onenote' | 'teams'>;
  title?: string;
  prompt: string;
  /** The document before the prompt; shape depends on the surface (see scenarios/README.md). */
  seed: Record<string, unknown>;
  /** What to answer if the planner asks a clarifying question. */
  clarification?: string;
  /** Approve plans and writes (default true). Set false to check that nothing is written unapproved. */
  approve?: boolean;
  expect?: ScenarioExpect;
  timeoutMs?: number;
}

export interface ScenarioResult {
  id: string;
  surface: string;
  prompt: string;
  passed: boolean;
  failures: string[];
  durationMs: number;
  status?: RunRecord['status'];
  applied: Record<string, number>;
  failedEffects: Array<{ kind: string; errorCode?: string }>;
  answer: string;
  errors: string[];
  gates: string[];
  steps: Array<{ kind: string; text: string }>;
  runs: RunRecord[];
  transcript: TranscriptEntry[];
  document: unknown;
}

type Sim = { snapshot(): unknown; restore(): void };

function installHost(scenario: Scenario): Sim {
  const seed = scenario.seed as never;
  switch (scenario.surface) {
    case 'excel':
      return installFakeExcel(excelSeed(seed));
    case 'word':
      return installFakeWord(wordSeed(seed));
    case 'powerpoint':
      return installFakePowerPoint(powerPointSeed(seed));
    case 'outlook':
      return installFakeOutlook(outlookSeed(seed));
  }
}

/** Plain JSON for the report (Maps become objects). */
function plain(value: unknown): unknown {
  return JSON.parse(
    JSON.stringify(value, (_k, v: unknown) => (v instanceof Map ? Object.fromEntries(v) : v)),
  );
}

const pending = (s: PanelState): boolean =>
  Boolean(
    s.pendingPlanClarification ||
    s.pendingCommandPlan ||
    s.pendingPlan ||
    s.pendingWrite ||
    s.pendingShare,
  );

export async function runScenario(scenario: Scenario): Promise<ScenarioResult> {
  const started = Date.now();
  const transcript: TranscriptEntry[] = [];
  const gates: string[] = [];
  const sim = installHost(scenario);
  const runtime = createApplicationRuntime();
  let dispose = (): void => {};
  try {
    const env = loadShellEnv();
    const bridge = selectBridge(scenario.surface);
    if (!bridge) throw new Error(`no bridge for ${scenario.surface}`);
    const { session } = await composeSession({
      config: shellConfigFromEnv(env),
      auth: scenarioAuth,
      bridge,
      unit: initialUnit({ surface: scenario.surface }),
      hooks: runtime.hooks,
      triggers: runtime.triggers,
      primeOnHostEvent: false,
      discoverCatalog: false,
      fetchImpl: liveFetch(env, transcript),
    });
    const controller = new PanelController(session, bridge);
    const panel = connectPanelRuntime({ session, bridge, controller, triggers: runtime.triggers });
    panel.start();
    dispose = () => {
      panel.dispose();
      controller.cancel();
      session.dispose();
    };
    const runsBefore = session.executions.list().length;

    // Act on every gate the pane would show, exactly as a user clicking through would.
    const approve = scenario.approve !== false;
    // Click like a person: on a later tick (never inside the state notification that staged the
    // gate), one gate at a time, then wait before treating the panel as idle.
    let lastClick = 0;
    let scheduled = false;
    const act = (): void => {
      scheduled = false;
      const s = controller.getState();
      if (s.pendingPlanClarification) {
        gates.push('clarification');
        controller.answerPlanClarification(scenario.clarification ?? 'Use sensible defaults.');
      } else if (s.pendingCommandPlan) {
        gates.push('confirm-plan');
        if (approve) controller.confirmCommandPlan();
        else controller.cancelCommandPlan();
      } else if (s.pendingPlan) {
        gates.push('approve-plan');
        if (approve) controller.approvePlan();
        else controller.rejectPlan();
      } else if (s.pendingWrite) {
        gates.push('approve-write');
        if (approve) controller.approvePendingWrite();
        else controller.rejectPendingWrite();
      } else if (s.pendingShare) {
        gates.push('approve-share');
        if (approve) controller.approvePendingShare();
        else controller.rejectPendingShare();
      } else return;
      lastClick = Date.now();
    };
    const unsubscribe = controller.subscribe((s) => {
      if (pending(s) && !scheduled) {
        scheduled = true;
        setTimeout(act, 300);
      }
    });

    const timeoutMs = scenario.timeoutMs ?? 240_000;
    await dispatchLikeComposer(controller, bridge, scenario);
    // Settle: idle (not busy, no gate) for 2 s in a row, or the time budget runs out.
    let idleSince = 0;
    while (Date.now() - started < timeoutMs) {
      const s = controller.getState();
      if (pending(s) && !scheduled) {
        scheduled = true; // a gate staged without a notification we caught
        setTimeout(act, 300);
      }
      if (!s.busy && !pending(s) && Date.now() - lastClick > 3000) {
        idleSince ||= Date.now();
        if (Date.now() - idleSince >= 2000) break;
      } else idleSince = 0;
      await new Promise((r) => setTimeout(r, 250));
    }
    unsubscribe();

    const state = controller.getState();
    const runs = session.executions.list().slice(runsBefore);
    const effects = runs.flatMap((r) => r.effects);
    const applied: Record<string, number> = {};
    for (const e of effects) if (e.ok) applied[e.kind] = (applied[e.kind] ?? 0) + 1;
    const answer = [...state.messages].reverse().find((m) => m.role === 'assistant')?.text ?? '';
    const errors = [
      ...(state.error ? [state.error] : []),
      ...state.steps.filter((s) => s.kind === 'error').map((s) => s.text),
    ];
    const document = plain(sim.snapshot());
    const result: ScenarioResult = {
      id: scenario.id,
      surface: scenario.surface,
      prompt: scenario.prompt,
      passed: false,
      failures: [],
      durationMs: Date.now() - started,
      ...(runs.length ? { status: runs.at(-1)!.status } : {}),
      applied,
      failedEffects: effects
        .filter((e) => !e.ok)
        .map((e) => ({ kind: e.kind, ...(e.errorCode ? { errorCode: e.errorCode } : {}) })),
      answer,
      errors: [...new Set(errors)],
      gates,
      steps: state.steps.map((s) => ({ kind: s.kind, text: s.text })),
      runs,
      transcript,
      document,
    };
    if (Date.now() - started >= timeoutMs) result.failures.push(`timed out after ${timeoutMs} ms`);
    result.failures.push(...check(scenario, result));
    result.passed = result.failures.length === 0;
    return result;
  } finally {
    dispose();
    runtime.dispose();
    sim.restore();
  }
}

/**
 * Route the prompt exactly as the pane's composer does on Enter (App `dispatch`): pasted CLI runs
 * directly, actions go to the planner first, anything else is chat. Same helpers, same order.
 */
async function dispatchLikeComposer(
  controller: PanelController,
  bridge: DocBridge,
  scenario: Scenario,
): Promise<void> {
  const allowedIntents = intentsForManifest(await bridge.getCapabilities());
  const palette = commandPaletteFor(scenario.surface, allowedIntents);
  const inv = parseComposerInput(scenario.prompt, { kind: 'selection' }, palette);
  const program = extractDirectCommandProgram(inv.raw);
  if (program) return controller.runDirectCommands(program);
  const intent = inferImplicitIntent(scenario.surface, allowedIntents, inv);
  const routed =
    intent !== undefined && intent !== inv.intent
      ? { ...inv, intent, raw: `/${intent} ${inv.raw}`.trim() }
      : inv;
  const seed = invocationToSeed(routed);
  const grounding = invocationToGrounding(inv);
  const planFirst =
    isActuating(intent) ||
    (intent === undefined && shouldUsePlannerForFreeText(allowedIntents, inv));
  if (planFirst) return controller.proposePlan(seed, grounding);
  if (isActuating(intent)) return controller.runCommands(seed, grounding);
  return controller.send(seed, grounding);
}

function check(scenario: Scenario, r: ScenarioResult): string[] {
  const e = scenario.expect ?? {};
  const out: string[] = [];
  const status = e.status ?? 'completed';
  if (r.status !== status) out.push(`status is ${r.status ?? 'none'}, expected ${status}`);
  if ((e.noErrors ?? true) && r.errors.length) out.push(`errors shown: ${r.errors.join(' | ')}`);
  for (const [kind, min] of Object.entries(e.applied ?? {})) {
    if ((r.applied[kind] ?? 0) < min)
      out.push(`applied ${r.applied[kind] ?? 0} × ${kind}, expected ≥ ${min}`);
  }
  for (const kind of e.notApplied ?? []) {
    if (r.applied[kind]) out.push(`${kind} was applied ${r.applied[kind]} time(s), expected none`);
  }
  const answer = r.answer.toLowerCase();
  for (const text of e.answerIncludes ?? []) {
    if (!answer.includes(text.toLowerCase())) out.push(`answer does not mention "${text}"`);
  }
  if (e.formatsInclude?.length) {
    const formats = Object.values(
      (r.document as { formats?: Record<string, Record<string, unknown>> }).formats ?? {},
    );
    for (const facet of e.formatsInclude) {
      if (!formats.some((f) => f[facet] !== undefined)) out.push(`no range was given ${facet}`);
    }
  }
  const doc = JSON.stringify(r.document).toLowerCase();
  for (const text of e.documentIncludes ?? []) {
    if (!doc.includes(text.toLowerCase())) out.push(`document does not contain "${text}"`);
  }
  return out;
}
