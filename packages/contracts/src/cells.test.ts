import { describe, expect, it } from 'vitest';
import { fillRange, shiftFormula } from './cells.js';

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

  it('never shifts a quoted sheet name, and refuses a fill that leaves the sheet', () => {
    expect(shiftFormula("='Q1'!A1+'FY 2024'!B2", 1, 0)).toBe("='Q1'!A2+'FY 2024'!B3");
    expect(() => shiftFormula('=XFD1', 0, 1)).toThrow(/off the worksheet/);
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
});
