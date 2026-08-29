/**
 * The shape a COBie file takes once it has been read, before any rule looks at it.
 *
 * Deliberately dumb: a workbook is sheets, a sheet is a header row and data rows,
 * and every cell is a trimmed string. Typing is a validation concern, not a
 * parsing one — `Elevation` holding "n/a" is a defect the checker must *report*,
 * so the parser must not choke on it or coerce it away.
 */

/**
 * A row's position in the source file, one-based and counting the header, so it
 * matches what the user sees in Excel's row gutter. Every finding carries one;
 * a defect nobody can locate is not worth reporting.
 */
export type SourceRow = number;

export interface CobieRow {
  /** Row number in the original spreadsheet, header included. */
  readonly sourceRow: SourceRow;
  /** Cell values by *normalised* header. Empty cells are absent, not ''. */
  readonly cells: Readonly<Record<string, string>>;
}

export interface ParsedSheet {
  /** Sheet name exactly as it appears in the file. */
  readonly name: string;
  /** Header cells in file order, original spelling preserved for reporting. */
  readonly headers: readonly string[];
  readonly rows: readonly CobieRow[];
}

export interface ParsedWorkbook {
  readonly fileName: string;
  readonly sheets: readonly ParsedSheet[];
  /**
   * Non-fatal problems hit while reading — a sheet that could not be decoded, a
   * file with no header row. Surfaced as findings so a half-read file reports
   * *why* it looks empty rather than silently scoring zero.
   */
  readonly readWarnings: readonly string[];
}

/**
 * COBie's empty-cell convention. The spreadsheet has no null, so the schema
 * says an unknown required value is written "n/a" — which means a cell reading
 * "n/a" is simultaneously well-formed and empty of information. The checker
 * treats these as absent for completeness purposes and reports them separately,
 * because "declared unknown" and "left blank" are different failures with
 * different owners.
 */
const PLACEHOLDER_VALUES: readonly string[] = [
  'n/a', 'n\\a', 'na', 'nil', 'none', 'null', 'tbc', 'tbd', 'unknown',
  'unset', 'undefined', 'not applicable', 'not known', 'xxx', '-', '--', '?'
];

export function isPlaceholder(value: string | undefined): boolean {
  if (value === undefined) { return false; }
  return PLACEHOLDER_VALUES.indexOf(value.trim().toLowerCase()) !== -1;
}

/** Blank, whitespace-only, or a placeholder. */
export function isEmptyOrPlaceholder(value: string | undefined): boolean {
  return value === undefined || value.trim() === '' || isPlaceholder(value);
}

/**
 * Splits a COBie list cell. COBie specifies comma separation; exporters also
 * emit semicolons, and Excel's own "wrap in quotes" habit leaves stray quoting.
 * Accepting all three is the difference between reporting a real broken
 * reference and reporting the separator.
 */
export function splitList(value: string | undefined): string[] {
  if (value === undefined) { return []; }
  return value
    .split(/[,;]/)
    .map((part) => part.trim().replace(/^["']|["']$/g, '').trim())
    .filter((part) => part !== '');
}
