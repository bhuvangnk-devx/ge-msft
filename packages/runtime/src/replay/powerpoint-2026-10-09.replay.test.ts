import { describe, expect, it, vi } from 'vitest';
import {
  CONFIRMED_PLAN_CLOSE,
  CONFIRMED_PLAN_OPEN,
  type ActuationRequest,
  type ActuationResult,
  type CapabilityManifest,
  type ContextRef,
  type ResolvedContext,
  type SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';

/**
 * PowerPoint defects reported from live use on 2026-10-09 (docs/POWERPOINT-DEFECTS-2026-10-09.md): a confirmed
 * "add a summary slide" plan ended with `done "The SIT Slide has been successfully imported and
 * verified."` and nothing added, and a later task's requests carried earlier tasks' write outcomes.
 */
const cmd = (program: string) => '```cmd\n' + program + '\n```';

const SUMMARY_PLAN = [
  'Execute this user-confirmed plan in the open Microsoft 365 surface.',
  CONFIRMED_PLAN_OPEN,
  'original_request: Add a summary slide with the three key findings: Revenue is 500; Profit is 200; Two pipeline rows have no owner.',
  'intent: draft',
  'surface: powerpoint',
  'step 1: add slide "Summary Findings" after the selected slide',
  'step 2: add the three findings as bullets',
  CONFIRMED_PLAN_CLOSE,
].join('\n');

class PowerPointLikeBridge implements DocBridge {
  readonly surface = 'powerpoint' as const;
  applied: ActuationRequest[] = [];
  failNext = false;
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'powerpoint',
      contextKinds: ['slide', 'shape', 'document'],
      reads: ['read'],
      actuations: [
        { kind: 'insert-slide', surface: 'powerpoint', title: 'Insert slide', reversible: true },
        { kind: 'add-shape', surface: 'powerpoint', title: 'Add shape', reversible: true },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([
      { id: 'pp:slide:257#0', kind: 'slide', surface: 'powerpoint', title: 'Slide 1' },
    ]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    if (this.failNext) {
      this.failNext = false;
      return Promise.resolve({
        ok: false,
        changeId: request.changeId,
        kind: request.kind,
        error: { code: 'no_text', message: 'add-shape shapeType=textBox needs text' },
      });
    }
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'slide:268#147487855',
    });
  }
}

function streamOf(pieces: string[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream({
    pull(c) {
      if (i < pieces.length) c.enqueue(enc.encode(pieces[i++]!));
      else c.close();
    },
  });
}

/** A streamAssist fetch that replays one scripted model reply per call and records each query. */
function scriptedFetch(turns: string[]) {
  const queries: string[] = [];
  let call = 0;
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    const body = JSON.parse(init?.body ?? '{}') as {
      query?: { text?: string; parts?: Array<{ text?: string }> };
    };
    queries.push(body.query?.text ?? (body.query?.parts ?? []).map((p) => p.text ?? '').join('\n'));
    const text = turns[Math.min(call, turns.length - 1)] ?? cmd('done');
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
  });
  return { fetch: fetchImpl as unknown as typeof fetch, queries };
}

function setup(turns: string[]) {
  const bridge = new PowerPointLikeBridge();
  const { fetch, queries } = scriptedFetch(turns);
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'powerpoint' as const } },
    context: { docState: false },
  });
  return { bridge, queries, session };
}

async function collect(gen: AsyncGenerator<SseEvent | CommandLoopEvent>) {
  const events: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) events.push(e);
  return events;
}

const doneTurns = (events: Array<SseEvent | CommandLoopEvent>) =>
  events.flatMap((e) => (e.type === 'done' && 'turn' in e ? [e.turn] : []));

describe('a confirmed change plan is not done before any write was tried', () => {
  it('refuses the first done with nothing changed, then completes with the slide', async () => {
    const { bridge, events, queries } = await (async () => {
      const s = setup([
        cmd('done "The SIT Slide has been successfully imported and verified."'),
        cmd(
          'slide "Summary Findings" "Revenue is 500" "Profit is 200" "Two pipeline rows have no owner"',
        ),
        cmd('done'),
      ]);
      const events = await collect(
        s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => true }),
      );
      return { ...s, events };
    })();

    expect(queries[1]).toContain('done refused: nothing has been changed yet');
    expect(bridge.applied.map((r) => r.kind)).toEqual(['insert-slide']);
    expect(doneTurns(events)).toEqual([3]);
  });

  it('refuses only once: a second done with nothing changed stands', async () => {
    const s = setup([cmd('done'), cmd('done')]);
    const events = await collect(s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => true }));

    expect(s.bridge.applied).toEqual([]);
    expect(doneTurns(events)).toEqual([2]);
  });

  it('accepts done after the user rejected the write', async () => {
    const s = setup([cmd('slide "Summary Findings" "Revenue is 500"'), cmd('done')]);
    await collect(s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => false }));

    // On this fork rejecting a plan ends the task at once (no further model turn), so the done
    // guard is never reached; nothing is written and no done is refused.
    expect(s.bridge.applied).toEqual([]);
    expect(s.queries).toHaveLength(1);
    expect(s.queries.join('\n')).not.toContain('done refused');
  });

  it('asks a confirmed analysis plan for its answer once instead of ending silently (2026-10-12)', async () => {
    const ANALYSIS_PLAN = [
      'Execute this user-confirmed plan in the open Microsoft 365 surface.',
      CONFIRMED_PLAN_OPEN,
      'original_request: Analisis data ini dan sebutkan tiga tren teratas',
      'intent: summarize',
      'surface: powerpoint',
      'step 1: read the deck',
      'step 2: highlight any unusual patterns',
      CONFIRMED_PLAN_CLOSE,
    ].join('\n');
    const s = setup([
      cmd('done'),
      cmd('done "Three trends: revenue up, costs flat, risks rising."'),
    ]);
    const events = await collect(s.session.runCommands(ANALYSIS_PLAN, { approvePlan: () => true }));

    expect(s.queries[1]).toContain('nothing was changed and no answer was given');
    expect(doneTurns(events)).toEqual([2]);
    expect(s.bridge.applied).toEqual([]);
  });

  it('asks for the answer only once: a second bare done ends the task', async () => {
    const plan = SUMMARY_PLAN.replace('intent: draft', 'intent: summarize')
      .replace(
        'step 1: add slide "Summary Findings" after the selected slide',
        'step 1: read the deck',
      )
      .replace('step 2: add the three findings as bullets', 'step 2: identify the trends');
    const s = setup([cmd('done'), cmd('done ""')]);
    const events = await collect(s.session.runCommands(plan, { approvePlan: () => true }));
    expect(doneTurns(events)).toEqual([2]);
  });

  it('a change plan shares one refusal: after the change refusal a bare done ends it', async () => {
    const s = setup([cmd('done'), cmd('done'), cmd('done')]);
    const events = await collect(s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => true }));
    expect(doneTurns(events)).toEqual([2]);
    expect(s.queries[1]).toContain('done refused: nothing has been changed yet');
  });

  it('leaves a task that is not a change plan alone', async () => {
    const s = setup([cmd('done')]);
    const events = await collect(
      s.session.runCommands('What is on slide 2?', { approvePlan: () => true }),
    );

    expect(doneTurns(events)).toEqual([1]);
  });
});

describe("a task's requests do not carry earlier tasks' write outcomes", () => {
  it('sends the second task no "Applied …" or "Attempted …" notes from the first', async () => {
    const s = setup([
      cmd('/add-shape slide=1 shapeType=rectangle left=560 top=20 width=140 height=60'),
      cmd('done'),
      cmd('slide "Summary Findings" "Revenue is 500"'),
      cmd('done'),
    ]);
    s.bridge.failNext = true;
    await collect(s.session.runCommands('Add a blue rectangle', { approvePlan: () => true }));
    await collect(s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => true }));

    const secondTask = s.queries.slice(2);
    expect(secondTask.length).toBeGreaterThan(0);
    for (const query of secondTask) {
      expect(query).not.toContain('Attempted add-shape');
      expect(query).not.toContain('Applied insert-slide');
    }
  });
});

describe('an unclosed finish when=verified program', () => {
  it('is told how to fix it, then completes with done', async () => {
    // Recorded live: the model ended its block after `finish when=verified` with no closing fence,
    // three turns running, and the task ended incomplete.
    const unclosed = '```cmd\nslide "Summary Findings" "Revenue is 500"\nfinish when=verified\n';
    const s = setup([unclosed, cmd('slide "Summary Findings" "Revenue is 500"\ndone')]);
    const events = await collect(s.session.runCommands(SUMMARY_PLAN, { approvePlan: () => true }));

    expect(s.queries[1]).toContain(
      'end the block with a line of three backticks, or end it with done',
    );
    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['insert-slide']);
    expect(doneTurns(events)).toEqual([2]);
  });
});
