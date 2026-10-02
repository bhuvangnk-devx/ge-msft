/**
 * Live wiring for scenario runs: the real Gemini Enterprise engine from `packages/web-shell/.env`,
 * signed in with the developer's `gcloud` account instead of the browser's Entra → WIF exchange.
 * Everything the pane would send still goes through the real client; only the token source and
 * the transcript recorder are swapped in. Test environments only.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { AuthClient } from '@ge/runtime';

export interface TranscriptEntry {
  at: string;
  url: string;
  status: number;
  request: string;
  /** The model's visible text for this call, joined from the streamed replies. */
  reply: string;
}

const REPO_ROOT = join(__dirname, '..', '..', '..', '..');

/** `packages/web-shell/.env`, then the process environment (same order as the Vite build). */
export function loadShellEnv(): Record<string, string> {
  const env: Record<string, string> = {};
  const file = join(REPO_ROOT, 'packages', 'web-shell', '.env');
  if (existsSync(file)) {
    for (const raw of readFileSync(file, 'utf8').split(/\r?\n/)) {
      const m = /^(?:export\s+)?([A-Z0-9_]+)\s*=\s*(.*)$/.exec(raw.trim());
      if (m) env[m[1]!] = m[2]!.replace(/^['"]|['"]$/g, '');
    }
  }
  for (const [key, value] of Object.entries(process.env)) {
    if (key.startsWith('VITE_') && value !== undefined) env[key] = value;
  }
  return env;
}

let cachedToken: { value: string; at: number } | undefined;

/** A Google access token from `gcloud` (or GE_ACCESS_TOKEN), refreshed every 45 minutes. */
export function googleAccessToken(): string {
  const inline = process.env.GE_ACCESS_TOKEN;
  if (inline) return inline;
  if (!cachedToken || Date.now() - cachedToken.at > 45 * 60_000) {
    const value = execFileSync('gcloud', ['auth', 'print-access-token'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }).trim();
    cachedToken = { value, at: Date.now() };
  }
  return cachedToken.value;
}

/** Stand-in for the signed-in Microsoft user; the Google token comes from {@link liveFetch}. */
export const scenarioAuth: AuthClient = {
  getIdToken: async () => 'scenario-runner-id-token',
  getIdentity: async () => ({ username: 'scenario-runner@test', displayName: 'Scenario runner' }),
};

/** Visible model text from a streamAssist JSON-array body (best effort, for the transcript). */
function replyText(body: string): string {
  try {
    const chunks = JSON.parse(body) as Array<{
      answer?: { replies?: Array<{ groundedContent?: { content?: { text?: string } } }> };
    }>;
    return chunks
      .flatMap((c) => c.answer?.replies ?? [])
      .map((r) => r.groundedContent?.content?.text ?? '')
      .join('');
  } catch {
    return body.slice(0, 2000);
  }
}

/**
 * The fetch the scenario session uses: answers the STS exchange with the gcloud token, adds the
 * quota project to Google API calls, and records each Gemini request/reply in `transcript`.
 */
export function liveFetch(
  env: Record<string, string>,
  transcript: TranscriptEntry[],
): typeof fetch {
  const quotaProject = env.VITE_WIF_USER_PROJECT || env.VITE_GCP_PROJECT;
  return (async (input: RequestInfo | URL, init: RequestInit = {}) => {
    const url = String(input instanceof Request ? input.url : input);
    if (url.startsWith('https://sts.googleapis.com/')) {
      return new Response(
        JSON.stringify({
          access_token: googleAccessToken(),
          token_type: 'Bearer',
          issued_token_type: 'urn:ietf:params:oauth:token-type:access_token',
          expires_in: 2700,
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    }
    const headers = new Headers(init.headers);
    if (quotaProject && url.includes('googleapis.com')) {
      headers.set('X-Goog-User-Project', quotaProject);
    }
    const res = await fetch(input, { ...init, headers });
    if (url.includes(':streamAssist')) {
      const body = await res.clone().text();
      transcript.push({
        at: new Date().toISOString(),
        url: url.replace(/projects\/[^/]+/, 'projects/…'),
        status: res.status,
        request: typeof init.body === 'string' ? init.body : '',
        reply: replyText(body),
      });
    }
    return res;
  }) as typeof fetch;
}
