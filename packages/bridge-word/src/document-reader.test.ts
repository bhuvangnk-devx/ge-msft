import { afterEach, describe, expect, it, vi } from 'vitest';
import { WordBridge } from './word-bridge.js';
import { OfficeWordHost } from './host-port.js';

afterEach(() => {
  vi.unstubAllGlobals();
});
describe('Word paged reads', () => {
  it('reads 6000 paragraphs including blanks without loading the whole body', async () => {
    const source = Array.from({ length: 6000 }, (_, i) => ({
      text: i % 40 === 0 ? '' : `paragraph ${i}`,
      styleBuiltIn: 'Normal',
    }));
    const loads: Array<{ skip: number; top: number }> = [];
    vi.stubGlobal('Word', {
      run: async (fn: (ctx: unknown) => Promise<unknown>) => {
        let ready = false;
        let options = { skip: 0, top: 0 };
        const paragraphs = {
          load(opts: { skip: number; top: number; select: string }) {
            expect(opts.select).toBe('text,styleBuiltIn');
            options = opts;
            loads.push(opts);
          },
          get items() {
            if (!ready) throw new Error('sync required');
            return source.slice(options.skip, options.skip + options.top);
          },
        };
        return fn({
          document: { body: { paragraphs } },
          sync: async () => {
            ready = true;
          },
        });
      },
    });
    const bridge = new WordBridge(new OfficeWordHost());
    const ids: string[] = [];
    let cursor: string | undefined;
    do {
      const page = await bridge.readDocumentPage({ cursor });
      ids.push(...page.units.map((u) => u.id));
      cursor = page.nextCursor;
    } while (cursor);
    expect(ids).toHaveLength(6000);
    expect(new Set(ids).size).toBe(6000);
    expect(ids.at(-1)).toBe('paragraph:6000');
    expect(loads.every((load) => load.top === 41)).toBe(true);
    expect(loads.at(-1)?.skip).toBe(5960);
  });
});
