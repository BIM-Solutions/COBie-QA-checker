/**
 * A CSV/TSV reader, because COBie is delivered as often in text as in xlsx —
 * one file per sheet, named `Component.csv` and so on.
 *
 * Written by hand rather than pulled in: the whole grammar is quoting and
 * newlines, the xlsx reader is already a 200KB lazy chunk, and a second parser
 * dependency to handle ten lines of logic is not a trade worth making.
 */

/** Splits on the first line terminator that appears, tolerating CRLF and CR. */
export function detectDelimiter(text: string): ',' | '\t' | ';' {
  const firstLine = text.split(/\r\n|\r|\n/, 1)[0] || '';
  const tabs = firstLine.split('\t').length - 1;
  const commas = firstLine.split(',').length - 1;
  const semis = firstLine.split(';').length - 1;
  if (tabs >= commas && tabs >= semis && tabs > 0) { return '\t'; }
  // Semicolons win only outright: a comma-delimited file whose cells contain
  // semicolon-separated lists (COBie's SpaceNames) must not be read as TSV.
  if (semis > commas && semis > 0) { return ';'; }
  return ',';
}

/**
 * Parses RFC 4180-ish CSV into a grid.
 *
 * Handles quoted fields, escaped quotes (`""`) and newlines inside quotes.
 * Does not handle a byte-order mark — `readCsv` in `workbookReader` strips that
 * before this ever sees the text, so the concern lives in one place.
 */
export function parseDelimited(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let index = 0;

  const endField = (): void => { row.push(field); field = ''; };
  const endRow = (): void => { endField(); rows.push(row); row = []; };

  while (index < text.length) {
    const char = text.charAt(index);

    if (inQuotes) {
      if (char === '"') {
        if (text.charAt(index + 1) === '"') { field += '"'; index += 2; continue; }
        inQuotes = false; index += 1; continue;
      }
      field += char; index += 1; continue;
    }

    if (char === '"' && field === '') { inQuotes = true; index += 1; continue; }
    if (char === delimiter) { endField(); index += 1; continue; }
    if (char === '\r') {
      // Swallow the LF of a CRLF pair so it does not open an empty row.
      endRow();
      index += text.charAt(index + 1) === '\n' ? 2 : 1;
      continue;
    }
    if (char === '\n') { endRow(); index += 1; continue; }

    field += char; index += 1;
  }

  // A file ending without a terminator still has a last row; one ending *with*
  // a terminator does not, and appending an empty row there would report a
  // phantom blank record on every well-formed file.
  if (field !== '' || row.length > 0) { endRow(); }

  return rows;
}
