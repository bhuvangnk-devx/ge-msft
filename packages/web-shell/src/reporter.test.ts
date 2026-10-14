import { describe, it, expect, vi } from 'vitest';
import {
  classifyFailure,
  fitBody,
  Reporter,
  type CopyDiagnostics,
  type Diagnostic,
  type DiagnosticError,
  type DiagnosticRun,
} from './reporter.js';
import { readBrowser, readHost, type ReportHost } from './host.js';

const SINCE = '2026-10-14T00:46:00.000Z';
const BEFORE = '2026-10-14T00:40:00.000Z';
const AT = '2026-10-14T00:46:10.000Z';

function diag(errors: DiagnosticError[] = [], runs: DiagnosticRun[] = []): Diagnostic {
  return { generatedAt: AT, surface: 'excel', routes: [], steps: [], errors, runs };
}
const err = (patch: Partial<DiagnosticError> = {}): DiagnosticError => ({
  at: AT,
  where: 'panel',
  message: 'Something failed',
  ...patch,
});
const run = (status: string, effects: DiagnosticRun['effects'] = []): DiagnosticRun => ({
  taskId: 't1',
  status,
  startedAt: SINCE,
  effects,
});

describe('classifyFailure', () => {
  it.each([
    ['policy block', diag([err({ code: 'policy_block' })], [run('blocked')]), 'blocked'],
    [
      'guard refused a write',
      diag(
        [err({ where: 'write', code: 'unsafe_formula' })],
        [run('incomplete', [{ ok: false, errorCode: 'unsafe_formula' }])],
      ),
      'blocked',
    ],
    [
      'write outcome unknown',
      diag([], [run('incomplete', [{ ok: false, errorCode: 'outcome_unknown' }])]),
      'uncertain',
    ],
    ['WIF exchange failed', diag([err({ code: 'auth_exchange' })], [run('completed')]), 'auth'],
    ['Gemini 403', diag([err({ code: 'http_403' })]), 'auth'],
    ['timeout', diag([err({ code: 'timeout' })], [run('completed')]), 'error'],
    [
      'Office write failed',
      diag(
        [err({ where: 'write', code: 'target_conflict' })],
        [run('incomplete', [{ ok: false, errorCode: 'target_conflict' }])],
      ),
      'error',
    ],
    ['run failed', diag([], [run('failed')]), 'error'],
    [
      'blocked run with an uncertain write (real report c2d92516, 2026-10-14)',
      diag(
        [
          err({ where: 'write', code: 'outcome_unknown' }),
          err({ where: 'write', code: 'prepare_failed' }),
        ],
        [
          run('blocked', [
            { ok: false, errorCode: 'outcome_unknown' },
            { ok: false, errorCode: 'prepare_failed' },
          ]),
        ],
      ),
      'uncertain',
    ],
    [
      'run started a moment before the pane turned busy',
      diag([], [{ ...run('failed'), startedAt: BEFORE, finishedAt: AT }]),
      'error',
    ],
    ['exhausted turns', diag([], [run('incomplete')]), 'error'],
  ] as const)('%s → %s', (_label, d, kind) => {
    expect(classifyFailure(d, SINCE)).toBe(kind);
  });

  it.each([
    ['nothing failed', diag([], [run('completed')])],
    ['cancelled', diag([err()], [run('cancelled')])],
    [
      'plan declined',
      diag(
        [err({ where: 'write', code: 'plan_unapproved' })],
        [run('incomplete', [{ ok: false, errorCode: 'plan_unapproved' }])],
      ),
    ],
    ['change declined', diag([], [run('incomplete', [{ ok: false, errorCode: 'unapproved' }])])],
    [
      'write retried, run completed',
      diag(
        [err({ where: 'write', code: 'anchor_drift' })],
        [run('completed', [{ ok: false, errorCode: 'anchor_drift' }, { ok: true }])],
      ),
    ],
    [
      'error from an earlier request',
      diag([err({ at: BEFORE })], [{ ...run('failed'), startedAt: BEFORE }]),
    ],
  ] as const)('%s → no report', (_label, d) => {
    expect(classifyFailure(d, SINCE)).toBeUndefined();
  });
});

const HOST: ReportHost = {
  label: 'Excel on the web',
  app: 'Excel',
  platform: 'OfficeOnline',
  browser: { name: 'Google Chrome', version: '153' },
};
const copy = (
  d: Diagnostic = diag([err({ code: 'timeout' })], [run('failed')]),
): CopyDiagnostics => ({
  question: 'Sum column B',
  diagnostic: d,
});
const answer = (status: number, body?: unknown, headers: Record<string, string> = {}): Response =>
  new Response(body === undefined ? null : JSON.stringify(body), { status, headers });

function reporter(
  responses: Array<Response | Error>,
  extra: { token?: (opts: { forceRefresh?: boolean }) => Promise<string> } = {},
) {
  const fetch = vi.fn(async (_url: string | URL | Request, _init?: RequestInit) => {
    const next = responses.shift();
    if (!next) throw new Error('no more responses');
    if (next instanceof Error) throw next;
    return next;
  });
  const sleep = vi.fn(async (_ms: number) => {});
  let ids = 0;
  let clock = Date.parse(AT);
  const r = new Reporter({
    url: 'https://reports.test/v1/reports',
    getToken: extra.token ?? (async () => 'entra-token'),
    fetch: fetch as unknown as typeof globalThis.fetch,
    sleep,
    now: () => clock,
    newId: () => `${++ids}`.padStart(32, '0'),
  });
  return { r, fetch, sleep, tick: (ms: number) => (clock += ms) };
}
const sentBody = (fetch: ReturnType<typeof reporter>['fetch'], i = 0) =>
  JSON.parse(fetch.mock.calls[i]![1]!.body as string);

describe('Reporter', () => {
  it('sends the Copy diagnostics content unchanged with a bearer token', async () => {
    const { r, fetch } = reporter([answer(202, { reportId: '0'.repeat(31) + '1' })]);
    const c = copy();
    await expect(r.report('error', HOST, c)).resolves.toEqual({
      status: 'sent',
      reportId: '0'.repeat(31) + '1',
    });
    const init = fetch.mock.calls[0]![1]!;
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer entra-token');
    expect(sentBody(fetch)).toEqual({
      reportId: '0'.repeat(31) + '1',
      kind: 'error',
      host: HOST,
      copyDiagnostics: c,
    });
  });

  it('retries 5xx after 5 s, 30 s and 2 min with the same reportId, then gives up', async () => {
    const { r, fetch, sleep } = reporter([answer(503), answer(503), answer(503), answer(503)]);
    await expect(r.report('error', HOST, copy())).resolves.toEqual({
      status: 'failed',
      reason: 'http_503',
    });
    expect(sleep.mock.calls.map(([ms]) => ms)).toEqual([5_000, 30_000, 120_000]);
    expect(new Set(fetch.mock.calls.map((_, i) => sentBody(fetch, i).reportId)).size).toBe(1);
  });

  it('treats a network error like a 5xx', async () => {
    const { r } = reporter([
      new TypeError('Failed to fetch'),
      answer(202, { reportId: 'x'.repeat(32) }),
    ]);
    await expect(r.report('error', HOST, copy())).resolves.toMatchObject({ status: 'sent' });
  });

  it('401 → one retry with a fresh token; 409 → one retry with a new reportId', async () => {
    const a = reporter([answer(401), answer(202, {})]);
    await expect(a.r.report('auth', HOST, copy())).resolves.toMatchObject({ status: 'sent' });
    expect(a.fetch).toHaveBeenCalledTimes(2);
    const b = reporter([answer(409), answer(202, {})]);
    await b.r.report('error', HOST, copy());
    expect(sentBody(b.fetch, 0).reportId).not.toBe(sentBody(b.fetch, 1).reportId);
  });

  it('429 waits Retry-After once; 400 and 403 stop', async () => {
    const a = reporter([answer(429, {}, { 'Retry-After': '42' }), answer(202, {})]);
    await expect(a.r.report('error', HOST, copy())).resolves.toMatchObject({ status: 'sent' });
    expect(a.sleep).toHaveBeenCalledWith(42_000);
    for (const status of [400, 403]) {
      const b = reporter([answer(status)]);
      await expect(b.r.report('error', HOST, copy())).resolves.toEqual({
        status: 'failed',
        reason: `http_${status}`,
      });
    }
  });

  it('413 drops the oldest half of the steps and resends', async () => {
    const d = diag([err({ code: 'timeout' })], [run('failed')]);
    d.steps = [{ kind: 'a' }, { kind: 'b' }, { kind: 'c' }, { kind: 'd' }];
    const { r, fetch } = reporter([answer(413), answer(202, {})]);
    await r.report('error', HOST, copy(d));
    expect(sentBody(fetch, 1).copyDiagnostics.diagnostic.steps).toEqual([
      { kind: 'c' },
      { kind: 'd' },
    ]);
  });

  it('asks for a fresh token on the 401 retry', async () => {
    const token = vi.fn(async (_opts: { forceRefresh?: boolean }) => 'entra-token');
    const { r } = reporter([answer(401), answer(202, {})], { token });
    await r.report('auth', HOST, copy());
    expect(token.mock.calls.map(([o]) => o.forceRefresh)).toEqual([false, true]);
  });

  it('removes tokens pasted into the question, and times out each request', async () => {
    const { r, fetch } = reporter([answer(202, {})]);
    await r.report('error', HOST, {
      ...copy(),
      question: 'why does eyJhbGciOi.eyJzdWIi.c2ln fail? ya29.abc',
    });
    expect(sentBody(fetch).copyDiagnostics.question).toBe('why does [token] fail? [token]');
    expect(fetch.mock.calls[0]![1]!.signal).toBeInstanceOf(AbortSignal);
  });

  it('never rejects', async () => {
    const { r } = reporter([]);
    const broken = { question: 'q', diagnostic: null } as unknown as CopyDiagnostics;
    await expect(r.report('error', HOST, broken)).resolves.toEqual({
      status: 'failed',
      reason: 'internal',
    });
  });

  it('no token → no request', async () => {
    const { r, fetch } = reporter([], {
      token: () => Promise.reject(new Error('interaction_required')),
    });
    await expect(r.report('error', HOST, copy())).resolves.toEqual({
      status: 'failed',
      reason: 'token',
    });
    expect(fetch).not.toHaveBeenCalled();
  });

  it('skips the same failure within 60 s, and anything past 10 in 10 minutes', async () => {
    const { r, fetch, tick } = reporter(Array.from({ length: 12 }, () => answer(202, {})));
    await r.report('error', HOST, copy());
    await expect(r.report('error', HOST, copy())).resolves.toEqual({
      status: 'skipped',
      reason: 'duplicate',
    });
    tick(61_000);
    await expect(r.report('error', HOST, copy())).resolves.toMatchObject({ status: 'sent' });
    for (let i = 0; i < 8; i++)
      await r.report('error', HOST, copy(diag([err({ message: `e${i}` })])));
    await expect(r.report('error', HOST, copy(diag([err({ message: 'e9' })])))).resolves.toEqual({
      status: 'skipped',
      reason: 'limit',
    });
    expect(fetch).toHaveBeenCalledTimes(10);
    tick(10 * 60_000 + 1);
    await expect(
      r.report('error', HOST, copy(diag([err({ message: 'later' })]))),
    ).resolves.toMatchObject({ status: 'sent' });
  });
});

describe('fitBody', () => {
  it('drops the oldest steps to fit 64 KB and never cuts the question', () => {
    const d = diag();
    d.steps = Array.from({ length: 400 }, (_, i) => ({
      kind: 'activity',
      text: `${i}`.padEnd(160, '.'),
    }));
    const question = 'q'.repeat(10_000);
    const fitted = fitBody({
      reportId: 'x'.repeat(32),
      kind: 'error',
      host: HOST,
      copyDiagnostics: { question, diagnostic: d },
    })!;
    expect(new TextEncoder().encode(JSON.stringify(fitted)).length).toBeLessThanOrEqual(65_536);
    expect(fitted.copyDiagnostics.question).toBe(question);
    expect(
      (fitted.copyDiagnostics.diagnostic.steps.at(-1) as { text: string }).text.startsWith('399'),
    ).toBe(true);
  });
});

describe('readHost', () => {
  const chrome = {
    brands: [
      { brand: 'Google Chrome', version: '153' },
      { brand: 'Not_A Brand', version: '8' },
      { brand: 'Chromium', version: '153' },
    ],
  };

  it('Word on the web (the values seen in the real pane, 2026-10-14)', async () => {
    const office = {
      context: { diagnostics: { host: 'Word', platform: 'OfficeOnline', version: '0.0.0.0' } },
    };
    await expect(readHost('word', office, { userAgentData: chrome })).resolves.toEqual({
      label: 'Word on the web',
      app: 'Word',
      platform: 'OfficeOnline',
      browser: { name: 'Google Chrome', version: '153' },
    });
  });

  it('desktop keeps the Office build; new Outlook is named by its client', async () => {
    const word = await readHost(
      'word',
      { context: { diagnostics: { host: 'Word', platform: 'PC', version: '16.0.17928.20114' } } },
      { userAgent: 'Mozilla/5.0 Chrome/141.0.0.0 Safari/537.36 Edg/141.0.3537.71' },
    );
    expect(word).toMatchObject({
      label: 'Word on Windows',
      version: '16.0.17928.20114',
      browser: { name: 'Microsoft Edge', version: '141.0.3537.71' },
    });
    const outlook = await readHost(
      'outlook',
      {
        context: {
          diagnostics: { host: 'Outlook', platform: 'OfficeOnline', version: '0.0.0.0' },
          mailbox: { diagnostics: { hostName: 'newOutlookWindows' } },
        },
      },
      {},
    );
    expect(outlook).toMatchObject({
      label: 'new Outlook on Windows',
      platform: 'OfficeOnline',
      outlookClient: 'newOutlookWindows',
    });
  });

  it('falls back to the surface when Office gives no diagnostics', async () => {
    await expect(readHost('excel', undefined, {})).resolves.toMatchObject({
      label: 'Excel',
      app: 'Excel',
      platform: 'unknown',
      browser: { name: 'unknown', version: '' },
    });
  });
});

describe('readBrowser', () => {
  it.each([
    [
      'exact Edge build',
      {
        userAgentData: {
          brands: [
            { brand: 'Chromium', version: '141' },
            { brand: 'Microsoft Edge', version: '141' },
          ],
          getHighEntropyValues: async () => ({
            fullVersionList: [
              { brand: 'Chromium', version: '141.0.7390.55' },
              { brand: 'Microsoft Edge', version: '141.0.3537.71' },
            ],
          }),
        },
      },
      { name: 'Microsoft Edge', version: '141.0.3537.71' },
    ],
    [
      'exact build refused',
      {
        userAgentData: {
          brands: [{ brand: 'Microsoft Edge', version: '141' }],
          getHighEntropyValues: async () => {
            throw new Error('NotAllowedError');
          },
        },
      },
      { name: 'Microsoft Edge', version: '141' },
    ],
    [
      'Firefox',
      { userAgent: 'Mozilla/5.0 (Windows NT 10.0; rv:143.0) Gecko/20100101 Firefox/143.0' },
      { name: 'Firefox', version: '143.0' },
    ],
    [
      'Safari',
      {
        userAgent:
          'Mozilla/5.0 (Macintosh) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0 Safari/605.1.15',
      },
      { name: 'Safari', version: '26.0' },
    ],
    [
      'Mac web view',
      { userAgent: 'Mozilla/5.0 (Macintosh) AppleWebKit/605.1.15 (KHTML, like Gecko)' },
      { name: 'WebKit', version: '605.1.15' },
    ],
  ] as const)('%s', async (_label, nav, want) => {
    await expect(readBrowser(nav as never)).resolves.toEqual(want);
  });
});
