import { afterEach, describe, expect, it, vi } from 'vitest';
import { readPowerPointPage } from './document-reader.js';

afterEach(() => {
  vi.unstubAllGlobals();
});
describe('PowerPoint paged reads', () => {
  it('visits all 200 slides, pages a dense slide, and reports non-text shapes', async () => {
    const textLoads: number[] = [];
    const shapes = Array.from({ length: 200 }, (_, slide) =>
      Array.from({ length: slide === 199 ? 100 : 3 }, (_, shape) => ({
        id: String(shape),
        type: shape === 0 ? 'Image' : 'TextBox',
        textFrame: {
          textRange: {
            text: `slide ${slide + 1} shape ${shape}`,
            load: () => {
              textLoads.push(slide);
            },
          },
        },
      })),
    );
    vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
    vi.stubGlobal('PowerPoint', {
      run: async (fn: (ctx: unknown) => Promise<unknown>) =>
        fn({
          presentation: {
            slides: {
              load: () => undefined,
              items: shapes.map((items, i) => {
                let offset = 0;
                let count = 0;
                return {
                  id: `s${i}`,
                  shapes: {
                    load: (opts: { skip: number; top: number }) => {
                      offset = opts.skip;
                      count = opts.top;
                      expect(count).toBeLessThanOrEqual(41);
                    },
                    get items() {
                      return items.slice(offset, offset + count);
                    },
                  },
                };
              }),
            },
          },
          sync: async () => undefined,
        }),
    });
    const ids = new Set<string>();
    let cursor: string | undefined;
    let pages = 0;
    do {
      const page = await readPowerPointPage({ cursor });
      for (const unit of page.units) {
        expect(ids.has(unit.id)).toBe(false);
        ids.add(unit.id);
      }
      expect(page.limitations.join(' ')).toContain('speaker notes');
      cursor = page.nextCursor;
      pages++;
    } while (cursor);
    expect(pages).toBe(202);
    expect(ids.has('slide:200:s199/shape:99')).toBe(true);
    expect(textLoads.filter((i) => i === 199)).toHaveLength(99);
  });
});
