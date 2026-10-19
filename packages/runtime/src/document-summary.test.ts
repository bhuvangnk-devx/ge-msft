import { describe, expect, it } from 'vitest';
import { requestsDocumentSummary, type DocumentPage, type SseEvent } from '@ge/contracts';
import {
  documentFragments,
  DOCUMENT_REQUEST_BYTES,
  summarizeDocument,
  type DocumentModel,
} from './document-summary.js';

async function collect(events: AsyncIterable<SseEvent>): Promise<SseEvent[]> {
  const result: SseEvent[] = [];
  for await (const event of events) result.push(event);
  return result;
}
const page = (units: number, prefix = ''): DocumentPage => ({
  scope: 'Test document text',
  limitations: ['Images excluded.'],
  units: Array.from({ length: units }, (_, i) => ({
    id: `${prefix}${i}`,
    text: `Source ${prefix}${i}: ${'evidence '.repeat(150)}`,
  })),
});
function modelProbe(): { model: DocumentModel; prompts: string[] } {
  const prompts: string[] = [];
  return {
    prompts,
    model: async function* (prompt) {
      prompts.push(prompt);
      yield { type: 'token', text: 'Bounded summary with source references.' };
      yield { type: 'done' };
    },
  };
}

describe('complete document summaries', () => {
  it('preserves non-blocking provider policy outcomes alongside a completed summary', async () => {
    const model: DocumentModel = async function* () {
      yield { type: 'policy', verdict: 'allow', reason: 'Policy passed.' };
      yield { type: 'token', text: 'Summary.' };
      yield { type: 'done' };
    };
    const events = await collect(summarizeDocument('summary', async () => page(1), model));
    expect(events).toContainEqual({
      type: 'policy',
      verdict: 'allow',
      reason: 'Policy passed.',
    });
    expect(events.at(-1)).toEqual({ type: 'done' });
  });
  it('recognizes whole-source summaries without promoting a selection question', () => {
    for (const q of [
      'read the entire doc and summarise it',
      'summarize all 200 slides',
      'summary of this document',
      'Summarise the whole workbook',
    ])
      expect(requestsDocumentSummary(q)).toBe(true);
    expect(requestsDocumentSummary('summarize the selected range')).toBe(false);
  });

  it('reads 300 page-equivalents, includes the final evidence, and verifies every batch', async () => {
    const pages = Array.from({ length: 300 }, (_, index) => ({
      ...page(4, `p${index}:`),
      ...(index < 299 ? { nextCursor: String(index + 1) } : {}),
    }));
    const calls: number[] = [];
    const probe = modelProbe();
    const events = await collect(
      summarizeDocument(
        'Summarise the entire doc',
        async ({ cursor }) => {
          const index = Number(cursor ?? 0);
          calls.push(index);
          return pages[index]!;
        },
        probe.model,
      ),
    );
    expect(calls).toEqual([...pages.keys(), ...pages.keys()]);
    expect(probe.prompts.join('\n')).toContain('p299:3');
    expect(
      probe.prompts.every(
        (prompt) =>
          new TextEncoder().encode(JSON.stringify({ query: { text: prompt } })).length <=
          DOCUMENT_REQUEST_BYTES,
      ),
    ).toBe(true);
    expect(events.at(-1)).toEqual({ type: 'done' });
    const answers = events.filter((e) => e.type === 'token');
    expect(answers).toHaveLength(1);
    expect(answers[0]).toMatchObject({
      text: expect.stringContaining('1200 source units across 300 batches'),
    });
    expect(answers[0]).toMatchObject({ text: expect.stringContaining('Images excluded') });
  });

  it('does not lose huge Unicode source units during splitting', () => {
    const text = (
      Array.from({ length: 32 }, (_, i) => String.fromCharCode(i)).join('') + '\ud800😀汉字"\\\n'
    ).repeat(2000);
    const fragments = [...documentFragments({ id: 'huge', text })].map(
      (item) => JSON.parse(item) as { text: string; part: number },
    );
    expect(fragments.map((f) => f.text).join('')).toBe(text);
    expect(fragments.map((f) => f.part)).toEqual(fragments.map((_, i) => i));
  });

  it('publishes no answer if a later source batch changed', async () => {
    let reads = 0;
    const events = await collect(
      summarizeDocument(
        'summary',
        async () => {
          reads++;
          return page(1, reads === 1 ? 'original' : 'edited');
        },
        modelProbe().model,
      ),
    );
    expect(events.some((e) => e.type === 'token' || e.type === 'done')).toBe(false);
    expect(events.at(-1)).toMatchObject({
      type: 'error',
      message: expect.stringContaining('changed'),
    });
  });

  it('fails on repeating cursors instead of declaring coverage', async () => {
    const events = await collect(
      summarizeDocument(
        'summary',
        async ({ cursor }) => ({ ...page(1, cursor ?? 'first'), nextCursor: 'repeat' }),
        modelProbe().model,
      ),
    );
    expect(events.at(-1)).toMatchObject({
      type: 'error',
      message: expect.stringContaining('cursor'),
    });
  });

  it.each(['eof', 'block', 'error', 'commands', 'oversized'] as const)(
    'does not publish %s partial output',
    async (failure) => {
      const model: DocumentModel = async function* () {
        yield {
          type: 'token',
          text:
            failure === 'commands'
              ? '```cmd\nread\n```'
              : failure === 'oversized'
                ? 'x'.repeat(20_000)
                : 'partial',
        };
        if (failure === 'block')
          yield { type: 'policy', verdict: 'block', reason: 'Tenant policy' };
        if (failure === 'error') yield { type: 'error', code: 'network', message: 'Disconnected' };
        if (failure !== 'eof') yield { type: 'done' };
      };
      const events = await collect(summarizeDocument('summary', async () => page(1), model));
      expect(events.some((e) => e.type === 'token' || e.type === 'done')).toBe(false);
      expect(events.some((e) => e.type === 'policy' || e.type === 'error')).toBe(true);
    },
  );

  it('enforces task budgets and cancellation before more host or model calls', async () => {
    const probe = modelProbe();
    const events = await collect(
      summarizeDocument('summary', async () => page(100), probe.model, { maxCalls: 1 }),
    );
    expect(events.at(-1)).toMatchObject({
      type: 'error',
      message: expect.stringContaining('model calls'),
    });
    expect(probe.prompts).toHaveLength(1);
    const controller = new AbortController();
    controller.abort();
    await expect(
      collect(
        summarizeDocument(
          'summary',
          async () => {
            throw new Error('must not read');
          },
          probe.model,
          { signal: controller.signal },
        ),
      ),
    ).rejects.toThrow();
  });
});
