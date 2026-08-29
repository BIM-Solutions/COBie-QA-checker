import type { CobieRow, ParsedSheet, ParsedWorkbook } from '../../models/workbook';
import { normaliseHeader } from '../../models/cobieSchema';
import { detectDelimiter, parseDelimited } from './csv';

/**
 * Turns a file the user picked into a `ParsedWorkbook`.
 *
 * The xlsx reader is a dynamic import. It is by some distance the largest thing
 * this web part depends on, and a page that merely *hosts* the checker should
 * not pay for it — the chunk loads when someone actually runs a check. Keep it
 * that way: a static import here silently doubles the entry bundle.
 */

/** Cells arriving from the xlsx reader are typed; the checker wants strings. */
function cellToString(value: unknown): string {
  if (value === null || value === undefined) { return ''; }
  if (value instanceof Date) {
    // COBie dates are ISO 8601 without a zone. Excel hands back a Date built in
    // local time; formatting it via toISOString would shift it across midnight
    // for anyone west of UTC and report a date defect that is not in the file.
    return formatLocalIso(value);
  }
  if (typeof value === 'number') {
    return String(value);
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }
  return String(value);
}

function pad(value: number, width: number): string {
  let text = String(value);
  while (text.length < width) { text = '0' + text; }
  return text;
}

/** `YYYY-MM-DDTHH:MM:SS` in the date's own local fields, no zone applied. */
export function formatLocalIso(date: Date): string {
  return (
    pad(date.getFullYear(), 4) + '-' +
    pad(date.getMonth() + 1, 2) + '-' +
    pad(date.getDate(), 2) + 'T' +
    pad(date.getHours(), 2) + ':' +
    pad(date.getMinutes(), 2) + ':' +
    pad(date.getSeconds(), 2)
  );
}

/**
 * Builds a sheet from a raw grid.
 *
 * `headerOffset` is where the header row sits in the original file, one-based,
 * so `sourceRow` on every row points at the real Excel row. COBie templates
 * occasionally carry a title band above the headers; finding the header row
 * rather than assuming row 1 is what makes those files readable.
 */
export function sheetFromGrid(name: string, grid: string[][]): ParsedSheet {
  const headerIndex = findHeaderRow(grid);
  if (headerIndex === -1) {
    return { name, headers: [], rows: [] };
  }

  const headers = (grid[headerIndex] || []).map((header) => (header || '').trim());
  const rows: CobieRow[] = [];

  for (let i = headerIndex + 1; i < grid.length; i++) {
    const raw = grid[i] || [];
    const cells: Record<string, string> = {};
    let hasValue = false;

    for (let c = 0; c < headers.length; c++) {
      const header = headers[c];
      if (header === '') { continue; }
      const value = (raw[c] === undefined ? '' : String(raw[c])).trim();
      if (value === '') { continue; }
      // Keyed on the normalised header so rules can look a column up without
      // knowing which spelling this particular exporter used.
      cells[normaliseHeader(header)] = value;
      hasValue = true;
    }

    // Wholly blank rows are padding, not records. Reporting them would bury the
    // real findings under one "missing Name" per empty row — and COBie
    // templates ship with hundreds of them.
    if (!hasValue) { continue; }

    rows.push({ sourceRow: i + 1, cells });
  }

  return { name, headers, rows };
}

/**
 * The header row is the first row carrying a COBie key column. Falling back to
 * "first non-empty row" would pick up a title band and read the real headers as
 * data.
 *
 * Cell count is deliberately not part of the test. Requiring two filled cells
 * would reject a legitimate single-column sheet - a Floor list, a one-list
 * PickLists sheet - and reading it as headerless drops every row silently,
 * which is the worst way for a checker to be wrong.
 */
function findHeaderRow(grid: string[][]): number {
  const limit = Math.min(grid.length, 10);
  for (let i = 0; i < limit; i++) {
    const row = (grid[i] || []).map((cell) => normaliseHeader(String(cell || '')));
    if (row.indexOf('name') !== -1 || row.indexOf('email') !== -1) { return i; }
  }
  // No recognisable key column - a PickLists sheet, or a COBie profile with
  // renamed columns. Take the widest row in the window, earliest on a tie.
  //
  // Widest rather than first-with-content because a title band is one merged
  // cell and a header row is many, so width is what actually distinguishes
  // them; and earliest-on-a-tie because a single-column sheet has no width to
  // compare and its first row is its header.
  let best = -1;
  let widest = 0;
  for (let i = 0; i < limit; i++) {
    const filled = (grid[i] || []).filter((cell) => String(cell || '').trim() !== '').length;
    if (filled > widest) { widest = filled; best = i; }
  }
  return best;
}

async function readXlsx(file: ArrayBuffer, fileName: string): Promise<ParsedWorkbook> {
  // `/browser` rather than the bare package: the root has no CommonJS entry, and
  // the browser build is the one that reads an ArrayBuffer without pulling in
  // Node's fs shims. The dynamic import keeps it out of the entry bundle.
  const readXlsxFile = (await import(
    /* webpackChunkName: "xlsx-reader" */ 'read-excel-file/browser'
  )).default;

  const sheets: ParsedSheet[] = [];
  const readWarnings: string[] = [];

  try {
    // One call returns every sheet with its name attached, so there is no
    // second pass and no chance of the name and the data going out of step.
    const raw = await readXlsxFile(file);
    for (let i = 0; i < raw.length; i++) {
      const grid = raw[i].data.map((row) => row.map(cellToString));
      sheets.push(sheetFromGrid(raw[i].sheet, grid));
    }
  } catch (error) {
    // Becomes an error-severity finding rather than a thrown exception, so a
    // file the reader cannot open reports *why* instead of showing the user a
    // blank screen.
    readWarnings.push(
      `The file could not be read as a spreadsheet: ` +
      `${error instanceof Error ? error.message : String(error)}`
    );
  }

  return { fileName, sheets, readWarnings };
}

/** Strips a UTF-8 BOM, which otherwise corrupts the first header name. */
function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

function readCsv(text: string, fileName: string): ParsedWorkbook {
  const clean = stripBom(text);
  const grid = parseDelimited(clean, detectDelimiter(clean));
  // A CSV holds one sheet, and its name is the file name: `Component.csv`
  // is the Component sheet. Without that convention there is nothing to match
  // the schema against.
  const sheetName = fileName.replace(/\.[^.]+$/, '');
  return { fileName, sheets: [sheetFromGrid(sheetName, grid)], readWarnings: [] };
}

export interface FileSource {
  readonly name: string;
  arrayBuffer(): Promise<ArrayBuffer>;
  text(): Promise<string>;
}

/** Dispatches on extension. Anything not .csv/.tsv/.txt is tried as xlsx. */
export async function readWorkbook(file: FileSource): Promise<ParsedWorkbook> {
  const name = file.name;
  if (/\.(csv|tsv|txt)$/i.test(name)) {
    return readCsv(await file.text(), name);
  }
  return readXlsx(await file.arrayBuffer(), name);
}
