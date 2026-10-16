import { describe, it, expect } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import { replaceDraftOpening, withDraftGreeting } from './draft-greeting.js';

const CIMB = { replace: ['Yth.', 'Yth', 'Yang terhormat'], with: 'Halo' };

describe('replaceDraftOpening', () => {
  it.each([
    ['Yth. Pak Irpan,\n\nTerima kasih.', 'Halo Pak Irpan,\n\nTerima kasih.'],
    ['YTH Pak Irpan,', 'Halo Pak Irpan,'],
    ['Yth.Pak Irpan,', 'Halo Pak Irpan,'],
    ['Yang terhormat Bapak Adfizar,', 'Halo Bapak Adfizar,'],
    ['  Yth. Bu Sari,', '  Halo Bu Sari,'],
    ['<p>Yth. Pak Irpan,</p><p>Terima kasih.</p>', '<p>Halo Pak Irpan,</p><p>Terima kasih.</p>'],
    ['<div><p> Yth. Pak Irpan,</p></div>', '<div><p> Halo Pak Irpan,</p></div>'],
  ])('replaces the opening of %j', (body, expected) => {
    expect(replaceDraftOpening(body, CIMB)).toBe(expected);
  });

  it.each([
    'Halo Pak Irpan,',
    'Dear Rina,',
    'Terima kasih, Yth. tim.',
    '<p>Saya setuju.</p><p>Yth. Pak Irpan</p>',
    'Ythanks — not a greeting',
  ])('leaves %j as written', (body) => {
    expect(replaceDraftOpening(body, CIMB)).toBe(body);
  });
});

const req = (kind: ActuationRequest['kind'], params: Record<string, unknown>): ActuationRequest =>
  ({ changeId: asChangeId('c1'), surface: 'outlook', kind, params }) as ActuationRequest;

describe('withDraftGreeting', () => {
  it('changes the body of every Outlook email write', () => {
    for (const kind of ['reply-mail', 'create-mail', 'set-body', 'prepend-body'] as const) {
      const out = withDraftGreeting(
        req(kind, { mail: { body: 'Yth. Pak Budi,', subject: 'S' } }),
        CIMB,
      );
      expect((out.params as { mail: { body: string; subject: string } }).mail).toEqual({
        body: 'Halo Pak Budi,',
        subject: 'S',
      });
    }
    const viaText = withDraftGreeting(req('reply-mail', { text: 'Yth. Pak Budi,' }), CIMB);
    expect((viaText.params as { text: string }).text).toBe('Halo Pak Budi,');
  });

  it('leaves other writes, and everything without a brand greeting, unchanged', () => {
    const comment = req('add-comment', { text: 'Yth. Pak Budi,' });
    expect(withDraftGreeting(comment, CIMB)).toBe(comment);
    const reply = req('reply-mail', { mail: { body: 'Yth. Pak Budi,' } });
    expect(withDraftGreeting(reply, undefined)).toBe(reply);
  });
});
