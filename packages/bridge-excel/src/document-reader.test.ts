import { afterEach, describe, expect, it, vi } from 'vitest';
import { DOCUMENT_TILE_CELLS, readExcelPage } from './document-reader.js';

afterEach(() => {
  vi.unstubAllGlobals();
});
describe('Excel paged reads', () => {
  it.each([8, 70])(
    'covers 30 tabs × 1500 rows × %i columns without a full used-range value load',
    async (columns) => {
      let cells = 0;
      const rectangles = new Set<string>();
      vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
      vi.stubGlobal('Excel', {
        run: async (fn: (ctx: unknown) => Promise<unknown>) =>
          fn({
            workbook: {
              worksheets: {
                load: () => undefined,
                items: Array.from({ length: 30 }, (_, index) => ({
                  id: `s${index}`,
                  name: `Tab ${index}`,
                  visibility: index === 29 ? 'Hidden' : 'Visible',
                  getUsedRangeOrNullObject: () => ({
                    address: `'Tab ${index}'!C5:ZZ1504`,
                    rowCount: 1500,
                    columnCount: columns,
                    rowIndex: 4,
                    columnIndex: 2,
                    isNullObject: false,
                    load: (props: string) => {
                      expect(props).not.toContain('values');
                    },
                    get values(): never {
                      throw new Error('Unbounded values load');
                    },
                  }),
                  getRangeByIndexes: (row: number, col: number, height: number, width: number) => {
                    expect(height * width).toBeLessThanOrEqual(DOCUMENT_TILE_CELLS);
                    const key = `${index}:${row}:${col}`;
                    expect(rectangles.has(key)).toBe(false);
                    rectangles.add(key);
                    cells += height * width;
                    return {
                      address: `${key}:${height}:${width}`,
                      load: () => undefined,
                      values: Array.from({ length: height }, (_, i) =>
                        Array.from({ length: width }, (_, j) => (row + i) * columns + col + j),
                      ),
                      formulas: Array.from({ length: height }, () =>
                        Array.from({ length: width }, () => ''),
                      ),
                    };
                  },
                })),
              },
            },
            sync: async () => undefined,
          }),
      });
      let cursor: string | undefined;
      const unitIds = new Set<string>();
      let last = '';
      do {
        const page = await readExcelPage({ cursor });
        for (const unit of page.units) {
          expect(unitIds.has(unit.id)).toBe(false);
          unitIds.add(unit.id);
          last = unit.id;
        }
        cursor = page.nextCursor;
      } while (cursor);
      expect(cells).toBe(30 * 1500 * columns);
      expect(last).toContain('sheet:s29/row:1504/');
    },
    20_000,
  );
});
