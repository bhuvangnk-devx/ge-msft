/**
 * Error reports (docs/CNGPT-Reports-API.md): when a request fails, send the Copy diagnostics content
 * to the configured reports API with the user's Entra access token. The API reads the user from the
 * token; the body never names them. Nothing is stored: a report that cannot be sent is dropped.
 */
import type { ReportHost } from './host.js';
import { redactCredentials } from './controller.js';

export type ReportKind = 'blocked' | 'uncertain' | 'auth' | 'error';

/** The parts of a Copy diagnostics `diagnostic` the classifier reads. */
export interface DiagnosticError {
  at: string;
  where: string;
  kind?: string;
  code?: string;
  message: string;
}
export interface DiagnosticRun {
  taskId: string;
  status: string;
  startedAt: string;
  finishedAt?: string;
  effects: Array<{ ok: boolean; errorCode?: string }>;
}
export interface Diagnostic {
  generatedAt: string;
  surface: string;
  steps: unknown[];
  errors: DiagnosticError[];
  runs: DiagnosticRun[];
  [field: string]: unknown;
}
export interface CopyDiagnostics {
  question: string;
  diagnostic: Diagnostic;
}

const DECLINED = new Set(['unapproved', 'plan_unapproved']);
const BLOCKED = new Set(['policy_block', 'blocked', 'unsafe_url', 'unsafe_formula']);
const UNCERTAIN = new Set(['outcome_unknown', 'verification_incomplete']);
const AUTH = new Set(['auth_exchange', 'http_401', 'http_403']);

/**
 * Decide whether the request that started at `since` failed, and how. Undefined = no report:
 * nothing failed, the user cancelled, or the user declined the plan or the change.
 */
export function classifyFailure(diagnostic: Diagnostic, since: string): ReportKind | undefined {
  const errors = diagnostic.errors.filter((e) => e.at >= since && !DECLINED.has(e.code ?? ''));
  // `since` is stamped when the pane turns busy, a moment after the run starts: a run belongs to
  // this request if it finished (or is still running) after that.
  const run = diagnostic.runs
    .filter((r) => (r.finishedAt ?? r.startedAt) >= since || r.status === 'running')
    .at(-1);
  if (run?.status === 'cancelled') return undefined;
  const failedEffects = (run?.effects ?? []).filter(
    (e) => !e.ok && !DECLINED.has(e.errorCode ?? ''),
  );
  // `incomplete` with only declined effects is the user saying no, not a failure.
  const runFailed =
    run !== undefined &&
    (run.status === 'failed' ||
      run.status === 'blocked' ||
      (run.status === 'incomplete' && (failedEffects.length > 0 || run.effects.length === 0)));
  // A write error the model retried, in a run that then completed, is not a failure.
  if (!runFailed && !errors.some((e) => e.where !== 'write')) return undefined;
  const codes = new Set([
    ...errors.map((e) => e.code ?? ''),
    ...failedEffects.map((e) => e.errorCode ?? ''),
  ]);
  const has = (set: Set<string>): boolean => [...codes].some((c) => set.has(c));
  // A write that may have landed outranks everything: someone must check the document.
  if (has(UNCERTAIN)) return 'uncertain';
  if (run?.status === 'blocked' || has(BLOCKED)) return 'blocked';
  if (has(AUTH)) return 'auth';
  return 'error';
}

export interface ReportBody {
  reportId: string;
  kind: ReportKind;
  host: ReportHost;
  copyDiagnostics: CopyDiagnostics;
}

export type ReportResult =
  | { status: 'sent'; reportId: string }
  | { status: 'skipped'; reason: 'limit' | 'duplicate' | 'too_large' }
  | { status: 'failed'; reason: string };

export interface ReporterOptions {
  /** Full endpoint, e.g. `https://reports.example/v1/reports`. */
  url: string;
  /** An Entra access token for the reports API scope, silently (never a popup). */
  getToken(opts: { forceRefresh?: boolean }): Promise<string>;
  fetch?: typeof fetch;
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
  newId?: () => string;
}

const MAX_BYTES = 65_536;
const RETRY_DELAYS_MS = [5_000, 30_000, 120_000];
const WINDOW_MS = 10 * 60_000;
const MAX_PER_WINDOW = 10;
const DEDUP_MS = 60_000;
const MAX_RETRY_AFTER_MS = 2 * 60_000;
const REQUEST_TIMEOUT_MS = 15_000;

/** Fit the body under 64 KB by dropping the oldest steps, then the oldest routes and runs. */
export function fitBody(body: ReportBody): ReportBody | undefined {
  const b = structuredClone(body);
  const d = b.copyDiagnostics.diagnostic;
  const size = (): number => new TextEncoder().encode(JSON.stringify(b)).length;
  const lists = [d.steps, d.routes, d.runs].filter((l): l is unknown[] => Array.isArray(l));
  for (const list of lists) while (size() > MAX_BYTES && list.length > 0) list.shift();
  return size() <= MAX_BYTES ? b : undefined;
}

export class Reporter {
  private readonly sent: number[] = [];
  private readonly recent = new Map<string, number>();
  private readonly fetchImpl: typeof fetch;
  private readonly now: () => number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly newId: () => string;

  constructor(private readonly opts: ReporterOptions) {
    this.fetchImpl = opts.fetch ?? ((input, init) => globalThis.fetch(input, init));
    this.now = opts.now ?? Date.now;
    this.sleep = opts.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.newId = opts.newId ?? (() => crypto.randomUUID().replaceAll('-', ''));
  }

  /** Send one report. Resolves when it is stored, skipped or given up; never rejects. */
  async report(kind: ReportKind, host: ReportHost, copy: CopyDiagnostics): Promise<ReportResult> {
    try {
      // A user can paste a token into a question; it never leaves the pane.
      return await this.send(kind, host, { ...copy, question: redactCredentials(copy.question) });
    } catch {
      return { status: 'failed', reason: 'internal' };
    }
  }

  private async send(
    kind: ReportKind,
    host: ReportHost,
    copy: CopyDiagnostics,
  ): Promise<ReportResult> {
    const now = this.now();
    while (this.sent.length > 0 && now - this.sent[0]! > WINDOW_MS) this.sent.shift();
    if (this.sent.length >= MAX_PER_WINDOW) return { status: 'skipped', reason: 'limit' };
    const newest = copy.diagnostic.errors.at(-1);
    const fingerprint = [
      kind,
      copy.diagnostic.surface,
      newest?.code ?? '',
      newest?.message ?? '',
    ].join('|');
    const seen = this.recent.get(fingerprint);
    if (seen !== undefined && now - seen < DEDUP_MS)
      return { status: 'skipped', reason: 'duplicate' };
    let body = fitBody({ reportId: this.newId(), kind, host, copyDiagnostics: copy });
    if (!body) return { status: 'skipped', reason: 'too_large' };
    this.sent.push(now);
    this.recent.set(fingerprint, now);

    let retriedAuth = false;
    let retriedConflict = false;
    let retriedLimit = false;
    let delays = 0;
    for (;;) {
      let token: string;
      try {
        token = await this.opts.getToken({ forceRefresh: retriedAuth });
      } catch {
        // No token for the reports API (sign-in failed or no consent): nothing can be sent.
        return { status: 'failed', reason: 'token' };
      }
      let res: Response;
      try {
        res = await this.post(body, token);
      } catch {
        // Network down: back off like a 5xx.
        if (delays >= RETRY_DELAYS_MS.length) return { status: 'failed', reason: 'network' };
        await this.sleep(RETRY_DELAYS_MS[delays++]!);
        continue;
      }
      if (res.status === 202) {
        const answer = (await res.json().catch(() => ({}))) as { reportId?: unknown };
        const reportId = typeof answer.reportId === 'string' ? answer.reportId : body.reportId;
        return { status: 'sent', reportId };
      }
      if (res.status === 401 && !retriedAuth) {
        retriedAuth = true;
        continue;
      }
      if (res.status === 409 && !retriedConflict) {
        retriedConflict = true;
        body = { ...body, reportId: this.newId() };
        continue;
      }
      if (res.status === 413) {
        const steps = body.copyDiagnostics.diagnostic.steps;
        if (steps.length === 0) return { status: 'failed', reason: 'http_413' };
        body = structuredClone(body);
        body.copyDiagnostics.diagnostic.steps.splice(0, Math.ceil(steps.length / 2));
        continue;
      }
      if (res.status === 429 && !retriedLimit) {
        retriedLimit = true;
        const seconds = Number(res.headers.get('Retry-After'));
        await this.sleep(
          Math.min(
            Number.isFinite(seconds) && seconds > 0 ? seconds * 1000 : 60_000,
            MAX_RETRY_AFTER_MS,
          ),
        );
        continue;
      }
      if (res.status >= 500 && delays < RETRY_DELAYS_MS.length) {
        await this.sleep(RETRY_DELAYS_MS[delays++]!);
        continue;
      }
      return { status: 'failed', reason: `http_${res.status}` };
    }
  }

  private post(body: ReportBody, token: string): Promise<Response> {
    return this.fetchImpl(this.opts.url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
  }
}
