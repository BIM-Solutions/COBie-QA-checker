import type { Finding, Severity } from '../../models/findings';
import type { CobieSheetName, Requirement } from '../../models/cobieSchema';
import { BUILT_IN_PICK_LISTS, COBIE_SCHEMA, normaliseHeader } from '../../models/cobieSchema';
import type { CobieRow } from '../../models/workbook';
import { isEmptyOrPlaceholder, isPlaceholder, splitList } from '../../models/workbook';
import type { SheetIndex, WorkbookIndex } from './workbookIndex';
import { valueOf } from './workbookIndex';

/**
 * The COBie rules themselves.
 *
 * Each rule is a plain function from the index to findings, so a rule can be
 * read, tested and argued with on its own. `RuleEngine` only sequences them.
 *
 * Severity follows requirement, everywhere and without exception:
 * `required` -> error, `expected` -> warning, `optional` -> not reported. That
 * mapping lives in `severityFor` and nowhere else, so "what counts as a
 * failure" is one decision rather than thirty.
 */

function severityFor(requirement: Requirement): Severity | undefined {
  if (requirement === 'required') { return 'error'; }
  if (requirement === 'expected') { return 'warning'; }
  return undefined;
}

/** Sheets whose rows are identified by a triple rather than by Name alone. */
const NOT_UNIQUE_ON_KEY: readonly CobieSheetName[] = [
  'Attribute', 'Coordinate', 'Document', 'Impact'
];

// ---------------------------------------------------------------------------
// Structure
// ---------------------------------------------------------------------------

/** Sheets the schema wants that the file does not have, and sheets it does not know. */
export function checkStructure(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];
  const names = Object.keys(COBIE_SCHEMA) as CobieSheetName[];

  for (let i = 0; i < names.length; i++) {
    const definition = COBIE_SCHEMA[names[i]];
    const severity = severityFor(definition.requirement);
    if (severity === undefined) { continue; }

    const sheet = index.sheets.get(names[i]);
    if (!sheet) {
      findings.push({
        ruleId: 'structure.missingSheet',
        category: 'structure',
        severity,
        sheet: definition.name,
        message: `The ${definition.name} sheet is missing.`
      });
      continue;
    }

    if (sheet.parsed.rows.length === 0) {
      findings.push({
        ruleId: 'structure.emptySheet',
        category: 'structure',
        severity,
        sheet: definition.name,
        message: `The ${definition.name} sheet has no rows.`
      });
    }
  }

  for (let i = 0; i < index.unknownSheets.length; i++) {
    // Info, not a warning: extra sheets are how most projects carry their own
    // data alongside COBie, and calling that a defect trains people to ignore
    // the report.
    findings.push({
      ruleId: 'structure.unknownSheet',
      category: 'structure',
      severity: 'info',
      sheet: index.unknownSheets[i],
      message: `"${index.unknownSheets[i]}" is not a COBie sheet and was not checked.`
    });
  }

  return findings;
}

/** Schema columns absent from a sheet that is present. */
export function checkColumns(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    for (let i = 0; i < sheet.bindings.length; i++) {
      const binding = sheet.bindings[i];
      if (binding.header !== undefined) { continue; }
      const severity = severityFor(binding.column.requirement);
      if (severity === undefined) { continue; }

      findings.push({
        ruleId: 'structure.missingColumn',
        category: 'structure',
        severity,
        sheet: sheet.definition.name,
        column: binding.column.name,
        message: `The ${binding.column.name} column is missing from ${sheet.definition.name}.`
      });
    }
  });

  return findings;
}

// ---------------------------------------------------------------------------
// Completeness and placeholders
// ---------------------------------------------------------------------------

/**
 * Blank and placeholder cells in columns that should carry a value.
 *
 * These are two findings, not one, and the split is the point. A blank cell is
 * usually an exporter that was never configured; an "n/a" is a human who looked
 * at the field and declined to fill it. They go to different people, so the
 * report must not merge them.
 */
export function checkCompleteness(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    for (let b = 0; b < sheet.bindings.length; b++) {
      const binding = sheet.bindings[b];
      if (binding.header === undefined) { continue; }
      const severity = severityFor(binding.column.requirement);
      if (severity === undefined) { continue; }

      const header = binding.header;
      const label = binding.displayHeader || binding.column.name;

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        const row = sheet.parsed.rows[r];
        const value = row.cells[header];

        if (value === undefined || value.trim() === '') {
          findings.push({
            ruleId: 'completeness.missingValue',
            category: 'completeness',
            severity,
            sheet: sheet.definition.name,
            row: row.sourceRow,
            column: label,
            message: `${label} is empty.`
          });
          continue;
        }

        if (isPlaceholder(value)) {
          findings.push({
            ruleId: 'placeholder.value',
            category: 'placeholder',
            severity,
            sheet: sheet.definition.name,
            row: row.sourceRow,
            column: label,
            value,
            message: `${label} holds the placeholder "${value}" rather than a value.`
          });
        }
      }
    }
  });

  return findings;
}

// ---------------------------------------------------------------------------
// Uniqueness
// ---------------------------------------------------------------------------

/**
 * Duplicate keys within a sheet.
 *
 * Reported on the *second* and later occurrences only. Flagging all of them
 * would double-count a simple duplicate and, worse, leave nothing to indicate
 * which row is the original.
 */
export function checkUniqueness(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet, name) => {
    const keyHeader = sheet.keyHeader;
    if (keyHeader === undefined) { return; }

    // Attribute, Document, Coordinate and Impact repeat a Name by design - the
    // same attribute against many rows - so identity there is Name plus the row
    // it describes.
    const compound = NOT_UNIQUE_ON_KEY.indexOf(name) !== -1;
    const seen = new Map<string, number>();

    for (let r = 0; r < sheet.parsed.rows.length; r++) {
      const row = sheet.parsed.rows[r];
      const key = identityOf(sheet, row, keyHeader, compound);
      if (key === undefined) { continue; }

      const first = seen.get(key);
      if (first === undefined) { seen.set(key, row.sourceRow); continue; }

      const value = row.cells[keyHeader];
      findings.push({
        ruleId: 'uniqueness.duplicateKey',
        category: 'uniqueness',
        severity: 'error',
        sheet: sheet.definition.name,
        row: row.sourceRow,
        column: sheet.definition.key,
        value,
        message: compound
          ? `This row repeats ${sheet.definition.key} "${value}" against the same target as row ${first}.`
          : `${sheet.definition.key} "${value}" is already used at row ${first}.`
      });
    }
  });

  return findings;
}

function identityOf(
  sheet: SheetIndex,
  row: CobieRow,
  keyHeader: string,
  compound: boolean
): string | undefined {
  const key = row.cells[keyHeader];
  if (isEmptyOrPlaceholder(key)) { return undefined; }
  const base = key.trim().toLowerCase();
  if (!compound) { return base; }

  const sheetName = valueOf(sheet, row, 'SheetName') || '';
  const rowName = valueOf(sheet, row, 'RowName') || '';
  // NUL, written as an escape rather than a raw byte: a literal one makes git
  // treat this file as binary and grep refuse to search it. The separator has
  // to be a character that cannot occur in a cell, or the parts run together -
  // name "A" + sheet "B C" and name "A B" + sheet "C" are different rows that a
  // space-joined key would collapse into one, reporting a duplicate that is not
  // there.
  return [base, sheetName.trim().toLowerCase(), rowName.trim().toLowerCase()].join('\u0000');
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

/**
 * Fixed foreign keys - Space to Floor, Component to Type, Zone to Space and the rest.
 *
 * A reference into a sheet the file does not contain is reported once per
 * source sheet rather than once per row. On a file missing its Type sheet that
 * is the difference between one actionable finding and fifty thousand
 * identical ones.
 */
export function checkReferences(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    const references = sheet.definition.references || [];

    for (let i = 0; i < references.length; i++) {
      const reference = references[i];
      const severity = severityFor(reference.requirement);
      if (severity === undefined) { continue; }

      const binding = sheet.bindings.filter((b) => b.column.name === reference.column)[0];
      // A missing column is already reported by checkColumns; repeating it here
      // as a broken reference would double-count one defect.
      if (!binding || binding.header === undefined) { continue; }

      const label = binding.displayHeader || reference.column;
      const target = index.sheets.get(reference.targetSheet);

      if (!target) {
        findings.push({
          ruleId: 'reference.missingTargetSheet',
          category: 'reference',
          severity,
          sheet: sheet.definition.name,
          column: label,
          message:
            `${label} points at the ${reference.targetSheet} sheet, ` +
            `which the file does not contain.`
        });
        continue;
      }

      const keys = keySetFor(target, reference.targetColumn);
      const header = binding.header;

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        const row = sheet.parsed.rows[r];
        const raw = row.cells[header];
        // Blank and placeholder cells are the completeness rule's business.
        if (isEmptyOrPlaceholder(raw)) { continue; }

        const values = reference.list ? splitList(raw) : [raw.trim()];
        for (let v = 0; v < values.length; v++) {
          if (keys.has(values[v].toLowerCase())) { continue; }
          findings.push({
            ruleId: 'reference.broken',
            category: 'reference',
            severity,
            sheet: sheet.definition.name,
            row: row.sourceRow,
            column: label,
            value: values[v],
            message: `"${values[v]}" is not a ${reference.targetSheet} ${target.definition.key}.`
          });
        }
      }
    }
  });

  return findings;
}

function keySetFor(target: SheetIndex, targetColumn: string | undefined): ReadonlySet<string> {
  if (targetColumn === undefined || targetColumn === target.definition.key) { return target.keys; }

  const binding = target.bindings.filter((b) => b.column.name === targetColumn)[0];
  const keys = new Set<string>();
  if (!binding || binding.header === undefined) { return keys; }

  for (let i = 0; i < target.parsed.rows.length; i++) {
    const value = target.parsed.rows[i].cells[binding.header];
    if (!isEmptyOrPlaceholder(value)) { keys.add(value.trim().toLowerCase()); }
  }
  return keys;
}

/**
 * Polymorphic references - the `SheetName` + `RowName` pairs used by Attribute,
 * Document, Coordinate, Impact, Assembly, Connection and Issue.
 *
 * These carry most of a COBie file's real defects, because nothing in a
 * spreadsheet stops the two halves from disagreeing, and an Attribute pointing
 * at a row that does not exist is invisible until something tries to use it.
 */
export function checkDynamicReferences(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    const references = sheet.definition.dynamicReferences || [];

    for (let i = 0; i < references.length; i++) {
      const reference = references[i];
      const severity = severityFor(reference.requirement);
      if (severity === undefined) { continue; }

      const sheetBinding = sheet.bindings.filter((b) => b.column.name === reference.sheetColumn)[0];
      const rowBinding = sheet.bindings.filter((b) => b.column.name === reference.rowColumn)[0];
      if (!sheetBinding || sheetBinding.header === undefined) { continue; }
      if (!rowBinding || rowBinding.header === undefined) { continue; }

      const rowLabel = rowBinding.displayHeader || reference.rowColumn;
      const sheetLabel = sheetBinding.displayHeader || reference.sheetColumn;

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        const row = sheet.parsed.rows[r];
        const targetSheetName = row.cells[sheetBinding.header];
        const rawRowName = row.cells[rowBinding.header];

        if (isEmptyOrPlaceholder(targetSheetName) || isEmptyOrPlaceholder(rawRowName)) { continue; }

        const target = resolveSheet(index, targetSheetName);
        if (!target) {
          findings.push({
            ruleId: 'reference.unknownSheetName',
            category: 'reference',
            severity,
            sheet: sheet.definition.name,
            row: row.sourceRow,
            column: sheetLabel,
            value: targetSheetName,
            message:
              `${sheetLabel} names "${targetSheetName}", ` +
              `which is not a COBie sheet in this file.`
          });
          continue;
        }

        const values = reference.list ? splitList(rawRowName) : [rawRowName.trim()];
        for (let v = 0; v < values.length; v++) {
          if (target.keys.has(values[v].toLowerCase())) { continue; }
          findings.push({
            ruleId: 'reference.brokenDynamic',
            category: 'reference',
            severity,
            sheet: sheet.definition.name,
            row: row.sourceRow,
            column: rowLabel,
            value: values[v],
            message: `"${values[v]}" is not a row on the ${target.definition.name} sheet.`
          });
        }
      }
    }
  });

  return findings;
}

function resolveSheet(index: WorkbookIndex, name: string): SheetIndex | undefined {
  const wanted = name.trim().toLowerCase();
  let found: SheetIndex | undefined;
  index.sheets.forEach((sheet, key) => {
    if (key.toLowerCase() === wanted) { found = sheet; }
  });
  return found;
}

// ---------------------------------------------------------------------------
// Formats
// ---------------------------------------------------------------------------

// Deliberately permissive. The job is to catch "Level 1" in a number column and
// "01/02/2026" in a date column, not to adjudicate exotic but legal addresses.
const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// COBie 2.4 requires ISO 8601 without a zone: 2026-01-15T09:30:00.
const ISO_DATE_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/;
const NUMBER = /^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$/;

export function checkFormats(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    for (let b = 0; b < sheet.bindings.length; b++) {
      const binding = sheet.bindings[b];
      if (binding.header === undefined) { continue; }

      const type = binding.column.type;
      if (type === 'text' || type === 'list') { continue; }

      const header = binding.header;
      const label = binding.displayHeader || binding.column.name;
      // A malformed *optional* value is still worth a note - it costs nothing to
      // say so - but it must never fail a file on its own.
      const severity: Severity = binding.column.requirement === 'required' ? 'error' : 'warning';

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        const row = sheet.parsed.rows[r];
        const value = row.cells[header];
        if (isEmptyOrPlaceholder(value)) { continue; }

        const problem = formatProblem(type, value.trim());
        if (problem === undefined) { continue; }

        findings.push({
          ruleId: `format.${type}`,
          category: 'format',
          severity,
          sheet: sheet.definition.name,
          row: row.sourceRow,
          column: label,
          value,
          message: `${label} ${problem}`
        });
      }
    }
  });

  return findings;
}

function formatProblem(type: string, value: string): string | undefined {
  if (type === 'email') {
    return EMAIL.test(value) ? undefined : `is not an email address: "${value}".`;
  }
  if (type === 'isoDateTime') {
    if (!ISO_DATE_TIME.test(value)) {
      return `is not an ISO 8601 date and time (YYYY-MM-DDTHH:MM:SS): "${value}".`;
    }
    return isRealDate(value) ? undefined : `is not a real date: "${value}".`;
  }
  if (type === 'number' || type === 'integer') {
    if (!NUMBER.test(value)) { return `is not a number: "${value}".`; }
    if (type === 'integer' && value.indexOf('.') !== -1) {
      return `must be a whole number: "${value}".`;
    }
    return undefined;
  }
  return undefined;
}

/**
 * Rejects well-formed dates that do not exist - 2026-02-30 matches the pattern
 * and is not a day. Parsed field by field rather than through `Date`, whose
 * constructor silently rolls February 30th over into March.
 */
function isRealDate(value: string): boolean {
  const year = parseInt(value.substring(0, 4), 10);
  const month = parseInt(value.substring(5, 7), 10);
  const day = parseInt(value.substring(8, 10), 10);
  const hour = parseInt(value.substring(11, 13), 10);
  const minute = parseInt(value.substring(14, 16), 10);
  const second = parseInt(value.substring(17, 19), 10);

  if (month < 1 || month > 12) { return false; }
  if (hour > 23 || minute > 59 || second > 59) { return false; }

  const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  const lengths = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day >= 1 && day <= lengths[month - 1];
}

// ---------------------------------------------------------------------------
// Pick lists
// ---------------------------------------------------------------------------

/**
 * Values outside their permitted enumeration.
 *
 * A file that ships its own PickLists sheet is checked against that; only the
 * lists COBie fixes rather than leaves to the project (`AssetType`,
 * `DurationUnit`, ...) fall back to the built-ins. Checking a project's own
 * Category values against a hard-coded list would report a defect on every
 * correctly classified file.
 */
export function checkPickLists(index: WorkbookIndex): Finding[] {
  const findings: Finding[] = [];

  index.sheets.forEach((sheet) => {
    for (let b = 0; b < sheet.bindings.length; b++) {
      const binding = sheet.bindings[b];
      const listName = binding.column.pickList;
      if (binding.header === undefined || listName === undefined) { continue; }

      const permitted = permittedValues(index, listName);
      if (permitted === undefined) { continue; }

      const header = binding.header;
      const label = binding.displayHeader || binding.column.name;

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        const row = sheet.parsed.rows[r];
        const value = row.cells[header];
        if (isEmptyOrPlaceholder(value)) { continue; }
        if (permitted.has(value.trim().toLowerCase())) { continue; }

        findings.push({
          ruleId: 'pickList.value',
          category: 'pickList',
          // Always a warning. Classification systems drift between COBie
          // versions and a project may legitimately extend a list, so an
          // unrecognised value is a question to ask, not a failure to assert.
          severity: 'warning',
          sheet: sheet.definition.name,
          row: row.sourceRow,
          column: label,
          value,
          message: `"${value}" is not in the ${listName} pick list.`
        });
      }
    }
  });

  return findings;
}

function permittedValues(
  index: WorkbookIndex,
  listName: string
): ReadonlySet<string> | undefined {
  const fromFile = index.filePickLists.get(normaliseHeader(listName));
  if (fromFile !== undefined) { return fromFile; }

  const builtIn = BUILT_IN_PICK_LISTS[listName];
  if (builtIn === undefined) { return undefined; }
  return new Set(builtIn.map((value) => value.toLowerCase()));
}
