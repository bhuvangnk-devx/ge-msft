/**
 * One reader for the Excel address argument every command takes (`set`, `grid`, `spill`, `format`,
 * `table`, `chart`, `cf`, `read`). A sheet name with spaces reaches us in three spellings —
 * `'Q3 Sales'!A1` (Excel's), `"Q3 Sales!A1"` (the whole address in double quotes) and `Q3 Sales!A1`
 * (bare) — and each command used to read them differently: CIMB UAT 2026-10-17 saw a `read` that
 * kept the double quotes and matched nothing, a `set` split at the space, and a chart that took
 * the fragment `Branch` for a defined name. Every spelling now becomes Excel's `'Q3 Sales'!A1`.
 */

const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`;
/** A cell or cell range, a whole-column span (I:I) or a whole-row span (2:2). */
const A1 = String.raw`(?:${CELL}(?::${CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})`;
const END = String.raw`(?=$|[\s,])`;
/** Sheet spellings, tried in order: 'quoted', "double-quoted", bare without spaces, bare with. */
const AREA = [
  new RegExp(String.raw`^'((?:[^']|'')+)'!(${A1})${END}`),
  new RegExp(String.raw`^"([^"!]+)"!(${A1})${END}`),
  new RegExp(String.raw`^([^\s!,'"=]+)!(${A1})${END}`),
  new RegExp(String.raw`^()(${A1})${END}`),
  new RegExp(String.raw`^([^!,'"=\n]+?)!(${A1})${END}`),
];
const BARE_SHEET = /^[A-Za-z_][A-Za-z0-9_.]*$/;

/** `Sheet!A1`, quoting the sheet name only when Excel needs it ('Q3 Sales'!A1, 'O''Brien'!A1). */
function sheetRef(sheet: string, ref: string, quoted: boolean): string {
  if (!sheet) return ref;
  const name = quoted ? sheet.replace(/''/g, "'") : sheet.trim();
  return BARE_SHEET.test(name) ? `${name}!${ref}` : `'${name.replace(/'/g, "''")}'!${ref}`;
}

/** A comma-separated area list at the start of `s`, or undefined when `s` does not start with one. */
function scanAreas(s: string): { address: string; end: number } | undefined {
  const areas: string[] = [];
  let i = 0;
  for (;;) {
    let hit: { text: string; end: number } | undefined;
    for (const [n, re] of AREA.entries()) {
      const m = re.exec(s.slice(i));
      if (m) {
        hit = { text: sheetRef(m[1]!, m[2]!, n === 0), end: i + m[0].length };
        break;
      }
    }
    if (!hit) return areas.length ? { address: areas.join(','), end: i } : undefined;
    areas.push(hit.text);
    i = hit.end;
    const comma = /^\s*,\s*/.exec(s.slice(i));
    if (!comma || !AREA.some((re) => re.test(s.slice(i + comma[0].length)))) {
      return { address: areas.join(','), end: i };
    }
    i += comma[0].length;
  }
}

/**
 * The address argument at the start of `text` and the text after it. An A1 area list comes back in
 * Excel's spelling; anything else (a defined name, a table ref) is the first token, with one pair of
 * wrapping double quotes removed.
 */
export function scanAddress(text: string): { address: string; rest: string } | undefined {
  const s = text.trimStart();
  if (!s) return undefined;
  if (s.startsWith('"')) {
    const close = s.indexOf('"', 1);
    if (close > 0) {
      const inner = s.slice(1, close);
      const after = s.slice(close + 1);
      const areas = scanAreas(inner.trim());
      if (areas && inner.trim().slice(areas.end).trim() === '' && /^(?:$|\s)/.test(after))
        return { address: areas.address, rest: after };
      // `"Q3 Sales"!A1`: a double-quoted sheet name.
      const sheet = scanAreas(s);
      if (sheet) return { address: sheet.address, rest: s.slice(sheet.end) };
      if (/^(?:$|\s)/.test(after)) return { address: inner, rest: after };
    }
  }
  const areas = scanAreas(s);
  if (areas) return { address: areas.address, rest: s.slice(areas.end) };
  const m = /^'[^']*'\S*|^\S+/.exec(s)!;
  return { address: m[0], rest: s.slice(m[0].length) };
}

/** A whole selector (`read <selector>`) in Excel's spelling; text that is not an address is kept. */
export function normalizeSelector(selector: string): string {
  const scanned = scanAddress(selector);
  return scanned && scanned.rest.trim() === '' ? scanned.address : selector.trim();
}
