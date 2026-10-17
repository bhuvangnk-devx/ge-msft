import { describe, expect, it } from 'vitest';
import { normalizeSelector, scanAddress } from './address.js';
import { parseCommandLine, parseProgramBlock } from './command-grammar.js';

describe('scanAddress', () => {
  it("reads every spelling of a sheet name with spaces as Excel's 'Sheet'!A1", () => {
    for (const text of [
      `'Branch Performance'!I3`,
      `"Branch Performance!I3"`,
      `"Branch Performance"!I3`,
      `Branch Performance!I3`,
    ])
      expect(scanAddress(text)).toEqual({ address: `'Branch Performance'!I3`, rest: '' });
  });

  it('keeps simple addresses, names and table refs as they are', () => {
    expect(scanAddress('Sales!F2 =C2-D2')).toEqual({ address: 'Sales!F2', rest: ' =C2-D2' });
    expect(scanAddress('B16 Total')).toEqual({ address: 'B16', rest: ' Total' });
    expect(scanAddress('TotalCell 42')).toEqual({ address: 'TotalCell', rest: ' 42' });
    expect(scanAddress('Leads[Gender] x')).toEqual({ address: 'Leads[Gender]', rest: ' x' });
    expect(scanAddress(`'O''Brien'!A1`)).toEqual({ address: `'O''Brien'!A1`, rest: '' });
  });

  it('reads a multi-area list in any spelling, with or without spaces after commas', () => {
    const want = `'Branch Performance'!C1:C55,'Branch Performance'!I1:I55`;
    for (const text of [
      `Branch Performance!C1:C55,Branch Performance!I1:I55 title="x"`,
      `"Branch Performance!C1:C55, Branch Performance!I1:I55" title="x"`,
      `'Branch Performance'!C1:C55, 'Branch Performance'!I1:I55 title="x"`,
    ])
      expect(scanAddress(text)).toEqual({ address: want, rest: ' title="x"' });
  });

  it('reads whole-column and whole-row spans', () => {
    expect(normalizeSelector('Branch Performance!I:I')).toBe(`'Branch Performance'!I:I`);
    expect(normalizeSelector('2:2')).toBe('2:2');
  });
});

describe('commands read the same address in every spelling (CIMB UAT 2026-10-17)', () => {
  it('set, read, format, chart and spill', () => {
    // Scenario run: `set Campaign Leads!H1 "ID Nasabah"` split at the space and failed.
    expect(parseCommandLine('set Campaign Leads!H1 "ID Nasabah"')).toMatchObject({
      verb: 'set',
      cell: `'Campaign Leads'!H1`,
      value: 'ID Nasabah',
    });
    expect(parseCommandLine('read "Branch Performance!A1:I5"')).toEqual({
      verb: 'read',
      selector: `'Branch Performance'!A1:I5`,
    });
    expect(parseCommandLine('format "Campaign Leads!H2:H21" numberFormat="00000"')).toMatchObject({
      verb: 'format',
      range: `'Campaign Leads'!H2:H21`,
      props: { numberFormat: '00000' },
    });
    // Scenario run: the bare spelling made the chart read `Branch` as a defined name.
    expect(
      parseCommandLine(
        'chart line Branch Performance!A1:A55,Branch Performance!I1:I55 title="Trend" series=columns',
      ),
    ).toEqual({
      verb: 'chart',
      chartType: 'line',
      range: `'Branch Performance'!A1:A55,'Branch Performance'!I1:I55`,
      props: { title: 'Trend', series: 'columns' },
    });
    expect(parseCommandLine('spill Branch Performance!K1 = ($rows)')).toMatchObject({
      verb: 'spill',
      range: `'Branch Performance'!K1`,
    });
  });

  it('set reads `= "…"` as that one value (CIMB scenario 2026-10-17)', () => {
    // The model wrote the formula in quotes after `=`, and every cell got a text formula.
    expect(parseCommandLine(`set "Branch Performance"!I3:I55 = "=IF(A3=A2,(F3-F2)/F2,0)"`)).toEqual(
      { verb: 'set', cell: `'Branch Performance'!I3:I55`, value: '=IF(A3=A2,(F3-F2)/F2,0)' },
    );
    expect(parseCommandLine('set A1 = "Gender"')).toEqual({
      verb: 'set',
      cell: 'A1',
      value: 'Gender',
    });
    // A real formula that joins quoted text stays a formula.
    expect(parseCommandLine('set A1 = "a" & "b"')).toMatchObject({ value: '= "a" & "b"' });
  });

  it('a let-bound read keeps no double quotes in its selector', () => {
    // Scenario run: `read "Branch Performance!I3"` kept the quotes, matched nothing, and the copy
    // wrote 52 empty cells.
    const program = parseProgramBlock(
      '```cmd\nlet $formula = read "Branch Performance!I3"\nset "Branch Performance!I4:I55" $formula\n```',
      new Set<string>(),
    );
    expect(program.entries[0]).toMatchObject({
      pipeline: { source: { src: 'read', selector: `'Branch Performance'!I3` } },
    });
    expect(program.entries[1]).toMatchObject({ verb: 'set', cell: `'Branch Performance'!I4:I55` });
  });
});
