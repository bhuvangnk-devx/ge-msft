import type { ActuationRequest } from '@ge/contracts';

/**
 * A brand's email opening: drafts that START with one of `replace` start with `with` instead
 * (CIMB, 2026-10-16: "Halo", not "Yth."). Only the first words of the draft change; the rest of
 * the text, and the reply language, stay as the model wrote them.
 */
export interface DraftGreeting {
  readonly replace: readonly string[];
  readonly with: string;
}

/** The Outlook writes whose `params.mail.body` / `params.text` is an email the user will send. */
const EMAIL_WRITES: ReadonlySet<ActuationRequest['kind']> = new Set([
  'reply-mail',
  'create-mail',
  'set-body',
  'prepend-body',
]);

const escape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * `body` with a listed opening swapped for the brand's. Leading spaces and HTML tags (`<p>`) are
 * kept; a listed word later in the text, or one that only starts a longer word, is left alone.
 */
export function replaceDraftOpening(body: string, greeting: DraftGreeting): string {
  // A word ending in a letter must end there ("Ythanks" is not "Yth"); one ending in "." may run
  // straight into the name ("Yth.Pak"), which then gets a space after the brand's opening.
  const words = [...greeting.replace]
    .map((w) => w.trim())
    .filter(Boolean)
    .sort((a, b) => b.length - a.length)
    .map((w) => (/[\p{L}\p{N}]$/u.test(w) ? `${escape(w)}(?![\\p{L}\\p{N}])` : escape(w)));
  if (!words.length) return body;
  const opening = new RegExp(String.raw`^(\s*(?:<[^>]*>\s*)*)(?:${words.join('|')})`, 'iu');
  return body.replace(opening, (match: string, lead: string) => {
    const joined = /^[\p{L}\p{N}]/u.test(body.slice(match.length));
    return `${lead}${greeting.with}${joined ? ' ' : ''}`;
  });
}

/** The request with its email body's opening replaced; other writes are returned unchanged. */
export function withDraftGreeting(
  request: ActuationRequest,
  greeting: DraftGreeting | undefined,
): ActuationRequest {
  if (!greeting || !EMAIL_WRITES.has(request.kind)) return request;
  const params = request.params as { mail?: { body?: unknown }; text?: unknown };
  const mailBody = typeof params.mail?.body === 'string' ? params.mail.body : undefined;
  const text = typeof params.text === 'string' ? params.text : undefined;
  if (mailBody === undefined && text === undefined) return request;
  return {
    ...request,
    params: {
      ...request.params,
      ...(mailBody !== undefined
        ? { mail: { ...params.mail, body: replaceDraftOpening(mailBody, greeting) } }
        : {}),
      ...(text !== undefined ? { text: replaceDraftOpening(text, greeting) } : {}),
    },
  };
}
