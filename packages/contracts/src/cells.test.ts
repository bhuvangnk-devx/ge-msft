import { describe, expect, it } from 'vitest';
import { fillFormula, fillRange, rangeForGrid } from './cells.js';

describe('fillFormula', () => {
  it('moves relative references down like the fill handle (UAT 2026-10 MoM growth column)', () => {
    expect(fillFormula('=IF(A2=A1,(F2-F1)/F1,"")', 1, 0)).toBe('=IF(A3=A2,(F3-F2)/F2,"")');
    expect(fillFormula('=IF(A2=A1,(F2-F1)/F1,"")', 55, 0)).toBe('=IF(A57=A56,(F57-F56)/F56,"")');
  });

  it('keeps $-anchored parts fixed', () => {
    expect(fillFormula('=$A$1+A$1+$A1+A1', 2, 1)).toBe('=$A$1+B$1+$A3+B3');
  });

  it('leaves quoted text, sheet names, function names, and table references unchanged', () => {
    expect(fillFormula(`='Q1 2026'!B2&"A1"`, 1, 0)).toBe(`='Q1 2026'!B3&"A1"`);
    expect(fillFormula('=LOG10(B2)+SUM(Sheet2!B2:B9)+Table1[@[Col A1]]', 1, 0)).toBe(
      '=LOG10(B3)+SUM(Sheet2!B3:B10)+Table1[@[Col A1]]',
    );
  });

  it('moves whole-column and whole-row references only along their own axis', () => {
    expect(fillFormula('=F2/SUM(F:F)+SUM(1:1)', 1, 0)).toBe('=F3/SUM(F:F)+SUM(2:2)');
    expect(fillFormula('=F2/SUM(F:F)+SUM(1:1)', 0, 1)).toBe('=G2/SUM(G:G)+SUM(1:1)');
  });

  it('turns a reference moved off the sheet into #REF!', () => {
    expect(fillFormula('=A1048576*2', 1, 0)).toBe('=#REF!*2');
  });
});

describe('fillRange', () => {
  it('fills a formula over every cell of an A1 range', () => {
    const cells = fillRange("'Branch Performance'!R2:R4", '=F2-F1');
    expect(cells).toEqual([['=F2-F1'], ['=F3-F2'], ['=F4-F3']]);
    // The grid matches the destination, so the bridge accepts it.
    if (!('error' in cells))
      expect(rangeForGrid("'Branch Performance'!R2:R4", cells.length, 1)).toBe(
        "'Branch Performance'!R2:R4",
      );
  });

  it('keeps one cell and named ranges as one cell, and repeats a literal over a range', () => {
    expect(fillRange('Sales!F2', '=C2-D2')).toEqual([['=C2-D2']]);
    expect(fillRange('TotalCell', '42')).toEqual([['42']]);
    expect(fillRange('B2:C3', 'n/a')).toEqual([
      ['n/a', 'n/a'],
      ['n/a', 'n/a'],
    ]);
  });

  it('rejects a reversed range and a range over the cell budget', () => {
    expect(fillRange('R57:R2', '=F2')).toMatchObject({ error: expect.stringMatching(/top-left/) });
    expect(fillRange('A1:A100001', '=B1')).toMatchObject({
      error: expect.stringMatching(/budget/),
    });
  });
});
