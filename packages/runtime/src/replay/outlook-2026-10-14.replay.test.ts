import { describe, expect, it, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { TriggerRegistry } from '@ge/triggers';
import { AssistSession, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';
import { completedEffectsExtension, installRuntimeExtensions } from '../extensions.js';
import { RuntimeHooks } from '../hooks.js';

/**
 * Client report 2026-10-14 (Indonesian, read mode on a meeting invitation): "/draft Buat balasan
 * sopan yang menerima undangan rapat" ended in `repair_exhausted` after three model turns:
 *   1. `inspect ref="mail:<id>"` and 2. `inspect mail:<id>` — the read snapshot labels the email
 *      `mail:<id>`, but inspect only matched the bare item id, so both reads failed silently;
 *   3. a write program that failed validation — the third consecutive failing turn exhausted the
 *      repair budget before the model ever saw why its write was rejected.
 * The diagnostics did not record turn 3's failing line, so `/save-draft` (unsupported on Outlook,
 * yet listed in the skill references) stands in for it here.
 */
const cmd = (program: string) => '```cmd\n' + program + '\n```';
const ITEM_ID = 'AAMkAGY0NTg4ZGVmLWEwMWQtNDcyZC05MDMzLTgwYWE5M2ZhZDUzYwBGAAAAAACQe=';

class OutlookLikeBridge implements DocBridge {
  readonly surface = 'outlook' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'outlook',
      contextKinds: ['mail-item'],
      reads: ['read'],
      actuations: [
        { kind: 'reply-mail', surface: 'outlook', title: 'Reply', reversible: false },
        { kind: 'set-body', surface: 'outlook', title: 'Set body', reversible: true },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([
      {
        id: ITEM_ID,
        kind: 'mail-item',
        surface: 'outlook',
        title: 'Email: Undangan Rapat: Tinjauan Anggaran Q4',
        live: true,
      },
    ]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([
      {
        ref: {
          id: ITEM_ID,
          kind: 'mail-item',
          surface: 'outlook',
          title: 'Email: Undangan Rapat: Tinjauan Anggaran Q4',
        },
        value: { as: 'text', text: 'Halo Adele, saya ingin mengundang Anda rapat. Irpan' },
      },
    ]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
}

function setup(turns: string[]) {
  const bridge = new OutlookLikeBridge();
  const queries: string[] = [];
  let call = 0;
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    const body = JSON.parse(init?.body ?? '{}') as { query?: { text?: string } };
    queries.push(body.query?.text ?? '');
    const text = turns[Math.min(call++, turns.length - 1)] ?? cmd('done');
    const chunk = {
      sessionInfo: { session: 's' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    const enc = new TextEncoder().encode(JSON.stringify([chunk]));
    return new Response(
      new ReadableStream({
        start(c) {
          c.enqueue(enc);
          c.close();
        },
      }),
      { status: 200 },
    );
  });
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetchImpl as unknown as typeof fetch,
  );
  // The panel's completion guard, so a run ends with the status and message the user sees.
  const hooks = new RuntimeHooks();
  installRuntimeExtensions([completedEffectsExtension], { hooks, triggers: new TriggerRegistry() });
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'outlook' as const } },
    hooks,
    context: { docState: false },
  });
  return { bridge, queries, session };
}

/** The loop's events, plus the guard's block message when the run ends with one. */
async function collect(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const events: Array<SseEvent | CommandLoopEvent> = [];
  try {
    for await (const e of gen) events.push(e);
  } catch (error) {
    events.push({
      type: 'error',
      code: 'guard',
      message: error instanceof Error ? error.message : String(error),
    });
  }
  return events;
}

const errorMessages = (events: Array<SseEvent | CommandLoopEvent>) =>
  events.flatMap((e) => (e.type === 'error' ? [e.message] : []));

const errorCodes = (events: Array<SseEvent | CommandLoopEvent>) =>
  events.flatMap((e) => (e.type === 'error' && 'code' in e ? [e.code] : []));

const BAD_WRITE = cmd(
  '/save-draft\n' +
    '/set-body html="<p>Halo Pak Irpan,</p><p>Terima kasih atas undangan rapatnya.</p>"\n' +
    'done "Draft balasan penerimaan undangan rapat telah berhasil dibuat."',
);
const GOOD_WRITE = cmd(
  'mail "Halo Pak Irpan,\\n\\nTerima kasih atas undangannya. Saya akan hadir.\\n\\nSalam,\\nAdele"\n' +
    'done "Draf balasan siap ditinjau."',
);
const REQUEST = '/draft Buat balasan sopan yang menerima undangan rapat';

describe('client run 2026-10-14: accept a meeting invitation in Indonesian', () => {
  it('resolves the email by the mail: locator the read snapshot shows, bare or as ref="…"', async () => {
    const s = setup([
      cmd(`inspect ref="mail:${ITEM_ID}"`),
      cmd(`inspect mail:${ITEM_ID}`),
      GOOD_WRITE,
    ]);
    await collect(s.session.runCommands(REQUEST, { approvePlan: () => true }));

    const fedBack = s.queries.slice(1).join('\n');
    expect(fedBack).not.toContain('inspect could not resolve');
    expect(fedBack).toContain('saya ingin mengundang Anda rapat');
    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
  });

  it('gives a rejected write its own corrections even after two failed reads', async () => {
    // Reads of something that does not exist still fail; they must not use up the write's repairs.
    const s = setup([
      cmd('inspect mail:not-this-email'),
      cmd('inspect mail:nor-this-one'),
      BAD_WRITE,
      GOOD_WRITE,
    ]);
    const events = await collect(s.session.runCommands(REQUEST, { approvePlan: () => true }));

    expect(errorCodes(events)).not.toContain('repair_exhausted');
    // The model saw why turn 3 was rejected before it wrote turn 4.
    expect(s.queries[3]).toContain('save-draft');
    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
  });

  it('still stops a write that keeps failing after its corrections', async () => {
    const s = setup([BAD_WRITE, BAD_WRITE, BAD_WRITE, GOOD_WRITE]);
    const events = await collect(s.session.runCommands(REQUEST, { approvePlan: () => true }));

    expect(errorCodes(events)).toContain('repair_exhausted');
    expect(s.bridge.applied).toEqual([]);
  });
});

describe('the end of a run where nothing was applied', () => {
  it('ends cancelled, without an error, when the user declines the staged reply', async () => {
    const s = setup([GOOD_WRITE]);
    const events = await collect(s.session.runCommands(REQUEST, { approvePlan: () => false }));

    expect(s.bridge.applied).toEqual([]);
    expect(errorMessages(events).join('\n')).not.toContain('may already have applied');
    expect(s.session.executions.list().at(-1)?.status).toBe('cancelled');
  });

  it('says no changes were applied when the corrections run out before any write', async () => {
    const s = setup([BAD_WRITE, BAD_WRITE, BAD_WRITE]);
    const messages = errorMessages(
      await collect(s.session.runCommands(REQUEST, { approvePlan: () => true })),
    ).join('\n');

    expect(messages).toContain('No changes were applied.');
    expect(messages).not.toContain('may already have applied');
  });
});
