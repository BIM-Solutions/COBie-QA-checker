import {
  COBIE_SCHEMA, PICKLIST_SHEET, headerCandidates, isCobieSheet, normaliseHeader
} from '../../models/cobieSchema';
import type { CobieColumn, CobieSheet, CobieSheetName } from '../../models/cobieSchema';
import type { CobieRow, ParsedSheet, ParsedWorkbook } from '../../models/workbook';
import { isEmptyOrPlaceholder } from '../../models/workbook';

/**
 * The lookup layer every rule reads through.
 *
 * Built once per run, for two reasons. Referential rules are the expensive ones
 * — Component→Type on a 50,000-row file is 50,000 lookups — and doing them
 * against a prepared key set turns the check from quadratic into linear. And
 * matching a schema column to a real header (case, spacing, `ExtSystem` versus
 * `ExternalSystem`) is fiddly enough that having every rule redo it invites
 * them to disagree.
 */

export interface ColumnBinding {
  readonly column: CobieColumn;
  /** Normalised header actually found in the file. Absent when the column is missing. */
  readonly header?: string;
  /** Original spelling, for messages. */
  readonly displayHeader?: string;
}

export interface SheetIndex {
  readonly definition: CobieSheet;
  readonly parsed: ParsedSheet;
  readonly bindings: readonly ColumnBinding[];
  /** Normalised key-column header, when the key column is present. */
  readonly keyHeader?: string;
  /**
   * Every key value on the sheet, lowercased. Lowercased because COBie key
   * matching is case-insensitive in practice: a Component pointing at type
   * "AHU-01" when the Type sheet says "AHU-1" is a real defect, but one
   * pointing at "ahu-01" is an exporter casing artefact, not a broken link.
   */
  readonly keys: ReadonlySet<string>;
}

export interface WorkbookIndex {
  readonly workbook: ParsedWorkbook;
  /** Only sheets the schema knows about and the file actually contains. */
  readonly sheets: ReadonlyMap<CobieSheetName, SheetIndex>;
  /** Sheet names in the file that the schema does not define. */
  readonly unknownSheets: readonly string[];
  /** Pick lists read from the file's own PickLists sheet, if it has one. */
  readonly filePickLists: ReadonlyMap<string, ReadonlySet<string>>;
}

/** Case- and space-insensitive lookup of a schema column in a parsed sheet. */
function bindColumn(column: CobieColumn, headers: readonly string[]): ColumnBinding {
  const candidates = headerCandidates(column);
  for (let i = 0; i < headers.length; i++) {
    const normalised = normaliseHeader(headers[i]);
    if (candidates.indexOf(normalised) !== -1) {
      return { column, header: normalised, displayHeader: headers[i] };
    }
  }
  return { column };
}

/** Finds a sheet by name, case-insensitively; COBie files vary on casing. */
function findSheet(workbook: ParsedWorkbook, name: string): ParsedSheet | undefined {
  const wanted = name.toLowerCase();
  for (let i = 0; i < workbook.sheets.length; i++) {
    if (workbook.sheets[i].name.trim().toLowerCase() === wanted) { return workbook.sheets[i]; }
  }
  return undefined;
}

/**
 * Reads the file's own PickLists sheet.
 *
 * The COBie PickLists sheet is column-per-list: the header names the list and
 * the cells below it are the permitted values. A file that ships one is checked
 * against it in preference to the built-ins, because a project is entitled to
 * extend the enumerations and flagging its own categories would be noise.
 */
function readPickLists(workbook: ParsedWorkbook): Map<string, Set<string>> {
  const lists = new Map<string, Set<string>>();
  const sheet = findSheet(workbook, PICKLIST_SHEET);
  if (!sheet) { return lists; }

  for (let c = 0; c < sheet.headers.length; c++) {
    const header = sheet.headers[c].trim();
    if (header === '') { continue; }
    const key = normaliseHeader(header);
    const values = new Set<string>();
    for (let r = 0; r < sheet.rows.length; r++) {
      const value = sheet.rows[r].cells[key];
      if (value !== undefined && value.trim() !== '') { values.add(value.trim().toLowerCase()); }
    }
    if (values.size > 0) { lists.set(key, values); }
  }
  return lists;
}

function collectKeys(parsed: ParsedSheet, keyHeader: string | undefined): Set<string> {
  const keys = new Set<string>();
  if (keyHeader === undefined) { return keys; }
  for (let i = 0; i < parsed.rows.length; i++) {
    const value = parsed.rows[i].cells[keyHeader];
    // Placeholders are excluded deliberately: a Type row named "n/a" must not
    // satisfy a Component pointing at "n/a". Both are defects and both should
    // be reported, rather than one silently resolving the other.
    if (!isEmptyOrPlaceholder(value)) { keys.add(value.trim().toLowerCase()); }
  }
  return keys;
}

export function buildIndex(workbook: ParsedWorkbook): WorkbookIndex {
  const sheets = new Map<CobieSheetName, SheetIndex>();
  const unknownSheets: string[] = [];

  for (let i = 0; i < workbook.sheets.length; i++) {
    const parsed = workbook.sheets[i];
    const trimmed = parsed.name.trim();
    if (trimmed.toLowerCase() === PICKLIST_SHEET.toLowerCase()) { continue; }

    // Match the schema case-insensitively but keep the canonical name as the key.
    const canonical = (Object.keys(COBIE_SCHEMA) as CobieSheetName[]).filter(
      (name) => name.toLowerCase() === trimmed.toLowerCase()
    )[0];

    if (canonical === undefined || !isCobieSheet(canonical)) {
      if (trimmed !== '') { unknownSheets.push(parsed.name); }
      continue;
    }

    const definition = COBIE_SCHEMA[canonical];
    const bindings = definition.columns.map((column) => bindColumn(column, parsed.headers));
    const keyBinding = bindings.filter((b) => b.column.name === definition.key)[0];
    const keyHeader = keyBinding ? keyBinding.header : undefined;

    sheets.set(canonical, {
      definition,
      parsed,
      bindings,
      keyHeader,
      keys: collectKeys(parsed, keyHeader)
    });
  }

  return {
    workbook,
    sheets,
    unknownSheets,
    filePickLists: readPickLists(workbook)
  };
}

/** The value of a schema column in a row, or undefined if the column is absent. */
export function valueOf(
  index: SheetIndex,
  row: CobieRow,
  columnName: string
): string | undefined {
  const binding = index.bindings.filter((b) => b.column.name === columnName)[0];
  if (!binding || binding.header === undefined) { return undefined; }
  return row.cells[binding.header];
}

/** Header spelling to quote in a message: the file's own, falling back to the schema's. */
export function headerLabel(index: SheetIndex, columnName: string): string {
  const binding = index.bindings.filter((b) => b.column.name === columnName)[0];
  return (binding && binding.displayHeader) || columnName;
}
