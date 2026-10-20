import { describe, it, expect } from 'vitest';
import { formulasStayInRow } from './actuate-plan.js';

describe('formulasStayInRow — formulas a sort can move with their row', () => {
  const ok = (formula: string, row = 2) => formulasStayInRow([[formula]], row, false);

  it('allows formulas that use only cells of their own row', () => {
    for (const f of [
      '=E2/D2',
      '=$E2/$D2',
      '=ROUND(E2/D2,2)',
      '=IF(A2="x 9",B2*2,0)',
      '=SUM(C2:F2)',
      '=LOG10(B2)',
      '=B2*1E5',
      '=TRUE',
    ])
      expect(ok(f), f).toBe(true);
    expect(
      formulasStayInRow(
        [
          ['', '=B2*2'],
          ['', '=B3*2'],
        ],
        2,
        false,
      ),
    ).toBe(true);
  });

  it('refuses anything that points outside its row or depends on position', () => {
    for (const f of [
      '=E3/D2',
      '=SUM(C2:C7)',
      '=E$2',
      '=$E$2',
      '=Sheet2!A2',
      "='Q1 Sales'!A2",
      '=[@Deposits]',
      '=SUM(A:A)',
      '=SUM(2:2)',
      '=Rate*E2',
      '=OFFSET(A2,0,1)',
      '=INDIRECT("A2")',
      '=ROW()',
      '=INDEX(A2:C2,1)',
      '=RAND()*A2',
    ])
      expect(ok(f), f).toBe(false);
  });

  it('skips the header row when the range has headers', () => {
    expect(formulasStayInRow([['=SUM(B2:B9)'], ['=B2*2']], 1, true)).toBe(true);
    expect(formulasStayInRow([['=SUM(B2:B9)'], ['=B2*2']], 1, false)).toBe(false);
  });
});
