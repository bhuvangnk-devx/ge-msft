import { describe, it, expect } from 'vitest';
import { quoteSheetNames } from './command-grammar.js';

describe('quoteSheetNames', () => {
  const sheets = ['Branch Performance', 'Sales', "O'Brien Q1"];

  it('quotes a known sheet name with a space written bare before !', () => {
    expect(quoteSheetNames('/sort-range range=Branch Performance!A1:I25 key=x', sheets)).toBe(
      "/sort-range range='Branch Performance'!A1:I25 key=x",
    );
    expect(quoteSheetNames('set D2 =SUM(Branch Performance!F2:F25)', sheets)).toBe(
      "set D2 =SUM('Branch Performance'!F2:F25)",
    );
    expect(quoteSheetNames("set O'Brien Q1!A1 1", sheets)).toBe("set 'O''Brien Q1'!A1 1");
  });

  it('leaves quoted names, simple names, unknown names and quoted values alone', () => {
    for (const line of [
      "set 'Branch Performance'!A1 1",
      'set Sales!A1 1',
      'set Other Sheet!A1 1',
      'set Sales!A1 "see Branch Performance!A1"',
    ])
      expect(quoteSheetNames(line, sheets)).toBe(line);
  });
});
