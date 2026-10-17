import { describe, expect, it } from 'vitest';
import type { ActuationRequest } from './capability.js';
import type { CellSnapshot } from './analysis.js';
import { cellsMatchRequest, excelInputValue, fillRange, shiftFormula } from './cells.js';

describe('shiftFormula', () => {
  it('shifts relative references and keeps absolute ones', () => {
    expect(shiftFormula('=IF(A2=A1,(F2-F1)/F1,"")', 3, 0)).toBe('=IF(A5=A4,(F5-F4)/F4,"")');
    expect(shiftFormula('=C2*$B$1+B$1+$A2', 2, 1)).toBe('=D4*$B$1+C$1+$A4');
    expect(shiftFormula("=SUM('Branch Performance'!F2:F9)", 1, 0)).toBe(
      "=SUM('Branch Performance'!F3:F10)",
    );
  });

  it('leaves text in quotes, function names and whole-column references alone', () => {
    expect(shiftFormula('=IF(B2>0,"A1 ok",LOG10(B2))', 1, 0)).toBe('=IF(B3>0,"A1 ok",LOG10(B3))');
    expect(shiftFormula('=SUM(A:A)', 4, 0)).toBe('=SUM(A:A)');
  });

  it('never shifts a quoted sheet name, and turns a reference moved off the sheet into #REF!', () => {
    // Was a throw; PR #4 (Gautam) matches Excel's fill handle. fillRange still refuses the fill.
    expect(shiftFormula("='Q1'!A1+'FY 2024'!B2", 1, 0)).toBe("='Q1'!A2+'FY 2024'!B3");
    expect(shiftFormula('=XFD1', 0, 1)).toBe('=#REF!');
    expect(shiftFormula('=A1048576*2', 1, 0)).toBe('=#REF!*2');
  });

  // From PR #4 by Gautam Rana.
  it('keeps $-anchored parts fixed', () => {
    expect(shiftFormula('=$A$1+A$1+$A1+A1', 2, 1)).toBe('=$A$1+B$1+$A3+B3');
  });

  it('leaves quoted text, sheet names, function names, and table references unchanged', () => {
    expect(shiftFormula(`='Q1 2026'!B2&"A1"`, 1, 0)).toBe(`='Q1 2026'!B3&"A1"`);
    expect(shiftFormula('=LOG10(B2)+SUM(Sheet2!B2:B9)+Table1[@[Col A1]]', 1, 0)).toBe(
      '=LOG10(B3)+SUM(Sheet2!B3:B10)+Table1[@[Col A1]]',
    );
    expect(shiftFormula('=Table1[@[Q1]]*B2', 1, 1)).toBe('=Table1[@[Q1]]*C3');
  });

  it('moves whole-column and whole-row references only along their own axis', () => {
    expect(shiftFormula('=F2/SUM(F:F)+SUM(1:1)', 1, 0)).toBe('=F3/SUM(F:F)+SUM(2:2)');
    expect(shiftFormula('=F2/SUM(F:F)+SUM(1:1)', 0, 1)).toBe('=G2/SUM(G:G)+SUM(1:1)');
  });

  it('returns a value that is not a formula unchanged', () => {
    expect(shiftFormula('Total', 5, 0)).toBe('Total');
  });
});

describe('fillRange', () => {
  it('fills a column with one formula, shifted per row like Excel fill-down (CIMB UAT #7)', () => {
    expect(fillRange("'Branch Performance'!I2:I4", '=IF(A2=A1,(F2-F1)/F1,"")')).toEqual([
      ['=IF(A2=A1,(F2-F1)/F1,"")'],
      ['=IF(A3=A2,(F3-F2)/F2,"")'],
      ['=IF(A4=A3,(F4-F3)/F3,"")'],
    ]);
  });

  it('fills a block across columns too, and repeats a plain value', () => {
    expect(fillRange('B2:C3', '=A2')).toEqual([
      ['=A2', '=B2'],
      ['=A3', '=B3'],
    ]);
    expect(fillRange('B2:B3', 'n/a')).toEqual([['n/a'], ['n/a']]);
  });

  it('is undefined for a single cell, so set keeps writing one cell', () => {
    expect(fillRange('Sales!F2', '=C2-D2')).toBeUndefined();
    expect(fillRange('F2:F2', '=C2-D2')).toBeUndefined();
  });

  it('refuses a reversed range, a range over the cell budget, and a fill that leaves the sheet', () => {
    // From PR #4 by Gautam Rana; ours throws so the compile path returns a correction.
    expect(() => fillRange('B5:A1', '=F2')).toThrow(/top-left/);
    expect(() => fillRange('R57:R2', '=F2')).toThrow(/top-left/);
    expect(() => fillRange('A1:A100001', '=B1')).toThrow(/budget/);
    expect(() => fillRange('A1:B2', '=XFD1')).toThrow(/off the worksheet/);
    // A `#REF!` already in the text does not hide one the fill would add.
    expect(() => fillRange('A1:B2', '=IF(XFD1=0,"#REF!",1)')).toThrow(/off the worksheet/);
  });
});

describe('excelInputValue', () => {
  it('marks conversions that change meaning (CIMB UAT 2026-10-16)', () => {
    for (const [text, stored] of [
      ['081234567807', 81234567807],
      ['+6281234567801', 6281234567801],
      ['4111111111111111', 4111111111111110],
      ['2E5', 200000],
      ['02/06/2026', 46059], // read month-first: 6 February 2026
      // Month labels lose their text as dates (live 2026-10-17), so they stay text.
      ['Jan 2026', 46023],
      ['Juni 2026', 46174],
      ['2 Agustus 2026', 46236],
    ] as const)
      expect(excelInputValue(text)).toEqual({ stored, keepsMeaning: false });
    expect(excelInputValue('3/4').keepsMeaning).toBe(false);
  });

  it('keeps conversions that keep meaning', () => {
    for (const [text, stored] of [
      ['100', 100],
      ['-12.5', -12.5],
      ['50%', 0.5],
      ['$1,234.50', 1234.5],
      ['TRUE', true],
      ['2026-06-24', 46197],
    ] as const)
      expect(excelInputValue(text)).toEqual({ stored, keepsMeaning: true });
    expect(excelInputValue('08:00 AM').stored).toBeCloseTo(8 / 24, 12);
  });

  it('leaves text Excel does not convert as text', () => {
    for (const text of ['Thamrin', '14/06/2026', '+62 812-3456', 'A1', ''])
      expect(excelInputValue(text)).toEqual({ stored: text, keepsMeaning: true });
  });
});

describe('cellsMatchRequest', () => {
  const request = (cells: string[][]) =>
    ({ params: { target: { range: 'A1' }, cells } }) as unknown as ActuationRequest;
  const snapshot = (values: unknown[][]) => ({ values, formulas: [] }) as unknown as CellSnapshot;

  it("accepts Excel's own conversion of the typed text, and nothing else", () => {
    expect(
      cellsMatchRequest(snapshot([[100, 0.5, 46197]]), request([['100', '50%', '2026-06-24']])),
    ).toBe(true);
    expect(cellsMatchRequest(snapshot([[8 / 24]]), request([['08:00 AM']]))).toBe(true);
    // A changed meaning is a real mismatch.
    expect(cellsMatchRequest(snapshot([[81234567807]]), request([['081234567807']]))).toBe(false);
    expect(cellsMatchRequest(snapshot([['']]), request([['0']]))).toBe(false);
  });
});
