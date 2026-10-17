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
      const actual = String(snapshot.values[r]?.[c] ?? '');
      if (actualFormula.startsWith('=')) return false;
      if (actual === String(v ?? '') || typeof v !== 'string') return actual === String(v ?? '');
      // Excel's own conversion of the typed text (100 → 100, 50% → 0.5, an ISO date → its serial)
      // is the approved value, not a mismatch; a conversion that changes meaning stays one. Fractions
      // compare within float noise.
      const { stored, keepsMeaning } = excelInputValue(v);
      if (!keepsMeaning) return false;
      return typeof stored === 'number'
        ? actual !== '' && Math.abs(Number(actual) - stored) <= 1e-9 * Math.max(1, Math.abs(stored))
        : actual === String(stored);
    }),
  );
}

/** What Excel stores for typed text, and whether that value still means what was typed. */
export interface ExcelInput {
  /** The value `Range.values` returns afterwards, parsed as en-US Excel does. */
  stored: string | number | boolean;
  /**
   * False when the stored value means something else: a lost leading 0 or +, digits past Excel's
   * 15, an exponent read into a code, or a numeric date whose day/month order the locale decides.
   * Such text must be written as text.
   */
  keepsMeaning: boolean;
}

/** English and Bahasa month names and short forms → month number. */
const MONTHS: Record<string, number> = {
  jan: 1,
  january: 1,
  januari: 1,
  feb: 2,
  february: 2,
  februari: 2,
  mar: 3,
  march: 3,
  maret: 3,
  apr: 4,
  april: 4,
  may: 5,
  mei: 5,
  jun: 6,
  june: 6,
  juni: 6,
  jul: 7,
  july: 7,
  juli: 7,
  aug: 8,
  august: 8,
  agu: 8,
  ags: 8,
  agustus: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  okt: 10,
  october: 10,
  oktober: 10,
  nov: 11,
  november: 11,
  dec: 12,
  des: 12,
  december: 12,
  desember: 12,
};

/** Excel's day serial (1900 date system), or undefined for an invalid date. */
function dateSerial(year: number, month: number, day: number): number | undefined {
  const time = Date.UTC(year, month - 1, day);
  if (new Date(time).getUTCDate() !== day || month < 1 || month > 12) return undefined;
  return (time - Date.UTC(1899, 11, 30)) / 86_400_000;
}

const fullYear = (y: string): number =>
  y.length === 4 ? Number(y) : Number(y) + (Number(y) < 30 ? 2000 : 1900);
const monthOf = (name: string): number =>
  Object.hasOwn(MONTHS, name.toLowerCase()) ? MONTHS[name.toLowerCase()]! : 0;

/**
 * Model of how Excel converts text typed into a cell (CIMB UAT 2026-10-16: `081…` lost its 0,
 * `+62…` became 6.28123E+12, and `02/06/2026` became 6 February). Text Excel leaves alone, and the
 * forms not modelled here (fractions, locale decimals), come back as the text itself.
 */
export function excelInputValue(text: string): ExcelInput {
  const t = text.trim();
  const keep = (stored: string | number | boolean): ExcelInput => ({ stored, keepsMeaning: true });
  const change = (stored: string | number | boolean): ExcelInput => ({
    stored,
    keepsMeaning: false,
  });
  if (t === '') return keep(text);
  if (/^(true|false)$/i.test(t)) return keep(t.toLowerCase() === 'true');
  let m = /^([+-]?)(\d+(?:\.\d*)?|\.\d+)(?:e([+-]?\d+))?$/i.exec(t);
  if (m) {
    const stored = Number(Number(t).toPrecision(15));
    const digits = (m[2]!.replace('.', '').replace(/^0+/, '') || '0').length;
    // A + before bare digits is a phone number (+6281…), and an exponent on a whole number is a code
    // (2E5); a signed decimal (+1.5e3) is still a number.
    const plusDigits = m[1] === '+' && /^\d+$/.test(m[2]!) && m[3] === undefined;
    const code = m[3] !== undefined && !m[2]!.includes('.');
    const lost = plusDigits || /^0\d/.test(m[2]!) || code || digits > 15;
    return lost ? change(stored) : keep(stored);
  }
  // Currency and thousands separators (en-US): $1,234.50 → 1234.5.
  if (/^\$?[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$/.test(t) && /[$,]/.test(t))
    return keep(Number(t.replace(/[$,]/g, '')));
  if ((m = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))%$/.exec(t))) return keep(Number(m[1]) / 100);
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([ap])m)?$/i.exec(t))) {
    const half = m[4]?.toLowerCase();
    const h = half ? (Number(m[1]) % 12) + (half === 'p' ? 12 : 0) : Number(m[1]);
    const min = Number(m[2]);
    const sec = Number(m[3] ?? 0);
    if (h < 24 && min < 60 && sec < 60) return keep((h * 3600 + min * 60 + sec) / 86_400);
  }
  // A numeric date reads month-first here and day-first in other locales, so it never keeps meaning.
  if ((m = /^(\d{1,2})([/.-])(\d{1,2})(?:\2(\d{2}|\d{4}))?$/.exec(t))) {
    const year = m[4] ? fullYear(m[4]) : new Date().getUTCFullYear();
    const serial = dateSerial(year, Number(m[1]), Number(m[3]));
    if (serial !== undefined) return change(serial);
  }
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t))) {
    const serial = dateSerial(Number(m[1]), Number(m[2]), Number(m[3]));
    if (serial !== undefined) return keep(serial);
  }
  // Month names (English or Bahasa) become dates, which loses the label: a summary labelled
  // "Jan 2026" stopped matching the data's text (live 2026-10-17), so these stay text.
  // "Jan 2026", "Juni 2026", "Agu-26" (a 2-digit number after a month is a year), "Jan 5".
  if ((m = /^([a-z]{3,9})\.?[\s./,-]+(\d{1,4})$/i.exec(t)) && monthOf(m[1]!) > 0) {
    const n = m[2]!;
    const serial =
      n.length === 1
        ? dateSerial(new Date().getUTCFullYear(), monthOf(m[1]!), Number(n))
        : dateSerial(n.length === 2 ? 2000 + Number(n) : Number(n), monthOf(m[1]!), 1);
    if (serial !== undefined) return change(serial);
  }
  // "2 Jun 2026", "2 Agustus", "2-Jun-26".
  if (
    (m = /^(\d{1,2})[\s./-]+([a-z]{3,9})\.?(?:[\s./,-]+(\d{2}|\d{4}))?$/i.exec(t)) &&
    monthOf(m[2]!) > 0
  ) {
    const year = m[3] ? fullYear(m[3]) : new Date().getUTCFullYear();
    const serial = dateSerial(year, monthOf(m[2]!), Number(m[1]));
    if (serial !== undefined) return change(serial);
  }
  // "Jun 2, 2026".
  if ((m = /^([a-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/i.exec(t)) && monthOf(m[1]!) > 0) {
    const serial = dateSerial(Number(m[3]), monthOf(m[1]!), Number(m[2]));
    if (serial !== undefined) return change(serial);
  }
  return keep(text);
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

const MAX_COLUMN = 16384;
const MAX_ROW = 1048576;
// A reference followed by a word character, `.`, `(` (a function), `!` (a sheet name) or `[`
// (a table name) is not a reference.
const CELL_REF = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w.(![])/y;
const COLUMN_SPAN = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w.(![])/y;
const ROW_SPAN = /(\$?)(\d+):(\$?)(\d+)(?![\w.(![])/y;
const WORD = /[\w.\\]+/y;

function matchAt(pattern: RegExp, text: string, at: number): RegExpExecArray | null {
  pattern.lastIndex = at;
  return pattern.exec(text);
}

/** The index after a `"…"` string or `'…'` sheet name that starts at `start` (doubled quote = escape). */
function quotedEnd(text: string, start: number): number {
  const quote = text[start];
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] !== quote) continue;
    if (text[i + 1] !== quote) return i + 1;
    i++;
  }
  return text.length;
}

/** The index after a `[…]` table or workbook reference (nested; `'` escapes the next character). */
function bracketEnd(text: string, start: number): number {
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "'") i++;
    else if (text[i] === '[') depth++;
    else if (text[i] === ']' && --depth === 0) return i + 1;
  }
  return text.length;
}

/**
 * A formula moved `rows` down and `columns` right, as Excel adjusts it on fill or copy: relative
 * A1 references shift, `$`-anchored parts stay. Whole-column (`F:F`) and whole-row (`1:1`)
 * references move only along their own axis. Text in double quotes, quoted sheet names
 * (`'Q1'!A1`), function names (`LOG10(`) and bracketed table references (`Table1[@[Q1]]`) are
 * left alone. A reference moved off the sheet becomes `#REF!`, as in Excel. A value that is not a
 * formula is returned as is.
 */
export function shiftFormula(formula: string, rows: number, columns: number): string {
  if (!formula.startsWith('=')) return formula;
  const row = (anchor: string, n: number) => (anchor ? n : n + rows);
  const column = (anchor: string, n: number) => (anchor ? n : n + columns);
  const rowOk = (n: number) => n >= 1 && n <= MAX_ROW;
  const columnOk = (n: number) => n >= 1 && n <= MAX_COLUMN;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'" || ch === '[') {
      const end = ch === '[' ? bracketEnd(formula, i) : quotedEnd(formula, i);
      out += formula.slice(i, end);
      i = end;
      continue;
    }
    const cell = matchAt(CELL_REF, formula, i);
    if (cell && columnOk(columnIndex(cell[2]!)) && rowOk(Number(cell[4]))) {
      const [text, ca = '', letters = '', ra = '', digits = ''] = cell;
      const c = column(ca, columnIndex(letters));
      const r = row(ra, Number(digits));
      out += columnOk(c) && rowOk(r) ? `${ca}${columnLabel(c)}${ra}${r}` : '#REF!';
      i += text.length;
      continue;
    }
    const cols = matchAt(COLUMN_SPAN, formula, i);
    if (cols && columnOk(columnIndex(cols[2]!)) && columnOk(columnIndex(cols[4]!))) {
      const [text, a1 = '', from = '', a2 = '', to = ''] = cols;
      const c1 = column(a1, columnIndex(from));
      const c2 = column(a2, columnIndex(to));
      out +=
        columnOk(c1) && columnOk(c2) ? `${a1}${columnLabel(c1)}:${a2}${columnLabel(c2)}` : '#REF!';
      i += text.length;
      continue;
    }
    const spanned = matchAt(ROW_SPAN, formula, i);
    if (spanned && rowOk(Number(spanned[2])) && rowOk(Number(spanned[4]))) {
      const [text, a1 = '', from = '', a2 = '', to = ''] = spanned;
      const r1 = row(a1, Number(from));
      const r2 = row(a2, Number(to));
      out += rowOk(r1) && rowOk(r2) ? `${a1}${r1}:${a2}${r2}` : '#REF!';
      i += text.length;
      continue;
    }
    const word = matchAt(WORD, formula, i);
    const next = word ? i + word[0].length : i + 1;
    out += formula.slice(i, next);
    i = next;
  }
  return out;
}

/**
 * The grid for writing one value over a whole range (`set I2:I57 =IF(A2=A1,…)`), like Excel's
 * Ctrl+Enter: the formula as written goes to the first cell and shifts for every other cell; a
 * plain value repeats. Undefined for a single cell, so `set` keeps writing just that cell. Throws
 * for a reversed range, a range over the cell budget, or a fill that moves a reference off the sheet.
 */
export function fillRange(address: string, value: string): string[][] | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*):\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(
    address.slice(address.lastIndexOf('!') + 1),
  );
  if (!match) return undefined;
  const rows = Number(match[4]) - Number(match[2]) + 1;
  const columns = columnIndex(match[3]!) - columnIndex(match[1]!) + 1;
  if (rows < 1 || columns < 1)
    throw new Error('The range must go from its top-left cell to its bottom-right cell.');
  if (rows * columns === 1) return undefined;
  if (rows * columns > 100_000) throw new Error('Write dimensions exceed the cell budget.');
  const cells = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => shiftFormula(value, r, c)),
  );
  // Excel would write #REF! here; a fill the model chose should not break the formula, so refuse.
  // Counted, so a `#REF!` already in the formula (even inside quoted text) cannot hide a new one.
  const refs = (text: string) => text.split('#REF!').length - 1;
  if (cells.some((row) => row.some((v) => refs(v) > refs(value))))
    throw new Error('The fill moves a reference off the worksheet; choose a smaller range.');
  return cells;
}
