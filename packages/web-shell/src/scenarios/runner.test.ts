import { describe, expect, it } from 'vitest';
import { geminiRefusal } from './runner.js';

const call = (status: number, reply = '') => ({ at: '', url: '', status, request: '', reply });

describe('geminiRefusal (WS5 T3)', () => {
  it('names the refusal when no call got an answer', () => {
    expect(geminiRefusal([])).toBe('no Gemini call was made');
    expect(geminiRefusal([call(429), call(429)])).toBe('Gemini refused every call (HTTP 429)');
    expect(geminiRefusal([call(401), call(503)])).toBe('Gemini refused every call (HTTP 401, 503)');
    expect(geminiRefusal([call(200, '')])).toBe('Gemini answered nothing (HTTP 200, empty)');
  });

  it('is undefined once any call got an answer: the run is a real result', () => {
    expect(geminiRefusal([call(200, '```plan\nintent ask\n```'), call(429)])).toBeUndefined();
  });
});
