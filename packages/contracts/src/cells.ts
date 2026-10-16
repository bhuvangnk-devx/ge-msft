import type { ActuationRequest } from './capability.js';
import type { CellSnapshot, CellValue } from './analysis.js';

export function gridForRequest(request: ActuationRequest): CellValue[][] {
  return request.params.cellValues ?? request.params.cells ?? [];
}
export function formulasForRequest(request: ActuationRequest): string[][] {
  const values = gridForRequest(request);
  return values.map((row, r) =>
    row.map((v, c) =>
      request.params.cellValues
        ? (request.params.cellFormulas?.[r]?.[c] ?? '')
        : typeof v === 'string' && v.startsWith('=')
          ? v
          : '',
    ),
  );
}
/** Compare formulas as expressions and values without locale-dependent display formatting. */
export function cellsMatchRequest(snapshot: CellSnapshot, request: ActuationRequest): boolean {
  const values = gridForRequest(request);
  const formulas = formulasForRequest(request);
  if (snapshot.values.length !== values.length || snapshot.values[0]?.length !== values[0]?.length)
    return false;
  return values.every((row, r) =>
    row.every((v, c) => {
      const wantedFormula = formulas[r]?.[c] ?? '';
      const actualFormula = snapshot.formulas?.[r]?.[c] ?? '';
      if (wantedFormula.startsWith('=')) return actualFormula === wantedFormula;
      return (
        !actualFormula.startsWith('=') && String(snapshot.values[r]?.[c] ?? '') === String(v ?? '')
      );
    }),
  );
}
export function hasFormulaErrors(snapshot: CellSnapshot): boolean {
  return snapshot.values.some((row, r) =>
    row.some(
      (v, c) =>
        snapshot.formulas?.[r]?.[c]?.startsWith('=') &&
        /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|N\/A|NUM!|SPILL!|CALC!|NULL!)/.test(String(v)),
    ),
  );
}
export function rangeForGrid(address: string, rows: number, columns: number): string {
  if (
    !Number.isInteger(rows) ||
    !Number.isInteger(columns) ||
    rows < 1 ||
    columns < 1 ||
    rows * columns > 100_000
  )
    throw new Error('Write dimensions exceed the cell budget.');
  const bang = address.lastIndexOf('!');
  const prefix = bang < 0 ? '' : address.slice(0, bang + 1);
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?::\$?([A-Za-z]{1,3})\$?([1-9]\d*))?$/.exec(
    address.slice(bang + 1),
  );
  if (!match) throw new Error('Choose an explicit A1 destination.');
  const index = (s: string): number =>
    [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const firstColumn = index(match[1]!);
  const firstRow = Number(match[2]);
  const lastColumn = firstColumn + columns - 1;
  const lastRow = firstRow + rows - 1;
  if (lastColumn > 16384 || lastRow > 1048576)
    throw new Error('Write exceeds worksheet dimensions.');
  if (match[3] && (index(match[3]) !== lastColumn || Number(match[4]) !== lastRow))
    throw new Error('The destination range dimensions do not match the result.');
  const label = (n: number): string => {
    let s = '';
    while (n) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  };
  const origin = `${label(firstColumn)}${firstRow}`;
  return `${prefix}${origin}${rows === 1 && columns === 1 ? '' : `:${label(lastColumn)}${lastRow}`}`;
}

const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const columnLabel = (index: number): string => {
  let s = '';
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

/**
 * A formula moved `rows` down and `columns` right, as Excel adjusts it on fill or copy: relative
 * A1 references shift, `$`-anchored parts stay. Text in double quotes, function names (`LOG10(`) and
 * whole-column references (`A:A`) are left alone. A value that is not a formula is returned as is.
 */
export function shiftFormula(formula: string, rows: number, columns: number): string {
  if (!formula.startsWith('=')) return formula;
  return (
    formula
      // Text in double quotes and quoted sheet names ('Q1'!A1) are never references to shift.
      .split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/)
      .map((part, i) =>
        i % 2 === 1
          ? part
          : part.replace(
              /(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?![\w(!])/g,
              (ref, colAbs: string, col: string, rowAbs: string, row: string) => {
                const c = colAbs ? columnIndex(col) : columnIndex(col) + columns;
                const r = rowAbs ? Number(row) : Number(row) + rows;
                if (c < 1 || c > 16384 || r < 1 || r > 1048576)
                  throw new Error(
                    `Filling ${ref} moves it off the worksheet; choose a smaller range.`,
                  );
                return `${colAbs}${colAbs ? col : columnLabel(c)}${rowAbs}${r}`;
              },
            ),
      )
      .join('')
  );
}

/**
 * The grid for writing one value over a whole range (`set I2:I57 =IF(A2=A1,…)`), like Excel's
 * Ctrl+Enter: the formula as written goes to the first cell and shifts for every other cell; a
 * plain value repeats. Undefined for a single cell, so `set` keeps writing just that cell.
 */
export function fillRange(address: string, value: string): string[][] | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*):\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(
    address.slice(address.lastIndexOf('!') + 1),
  );
  if (!match) return undefined;
  const rows = Number(match[4]) - Number(match[2]) + 1;
  const columns = columnIndex(match[3]!) - columnIndex(match[1]!) + 1;
  if (rows < 1 || columns < 1 || rows * columns === 1) return undefined;
  if (rows * columns > 100_000) throw new Error('Write dimensions exceed the cell budget.');
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => shiftFormula(value, r, c)),
  );
}
