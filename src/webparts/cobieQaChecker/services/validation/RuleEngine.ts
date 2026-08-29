import type { CheckRun, Finding, SheetSummary } from '../../models/findings';
import { compareFindings } from '../../models/findings';
import { COBIE_SCHEMA } from '../../models/cobieSchema';
import type { CobieSheetName } from '../../models/cobieSchema';
import type { ParsedWorkbook } from '../../models/workbook';
import { isEmptyOrPlaceholder } from '../../models/workbook';
import { buildIndex } from './workbookIndex';
import type { WorkbookIndex } from './workbookIndex';
import {
  checkColumns,
  checkCompleteness,
  checkDynamicReferences,
  checkFormats,
  checkPickLists,
  checkReferences,
  checkStructure,
  checkUniqueness
} from './rules';

/**
 * Runs every rule over a parsed workbook and rolls the findings up.
 *
 * Synchronous and single-pass on purpose. A 50,000-row COBie file indexes and
 * checks in well under a second, and making it async would buy nothing but a
 * spinner; the UI yields once before calling in so the button can show a
 * pending state, which is the whole of the concurrency story here.
 */

export interface CheckOptions {
  /** Recorded on the run so a stored result says who produced it. */
  readonly checkedBy: string;
  /**
   * Cap on findings kept. A file missing its Contact sheet generates one
   * finding per row per email column, which is hundreds of thousands on a real
   * deliverable - enough to exhaust the tab rather than inform anyone. The
   * counts stay exact; only the retained list is truncated.
   */
  readonly maxFindings?: number;
}

const DEFAULT_MAX_FINDINGS = 5000;

export function runChecks(workbook: ParsedWorkbook, options: CheckOptions): CheckRun {
  const index = buildIndex(workbook);

  // Order matters only for readability of the unsorted list; findings are
  // sorted before they are returned.
  const findings: Finding[] = ([] as Finding[])
    .concat(readWarningFindings(workbook))
    .concat(checkStructure(index))
    .concat(checkColumns(index))
    .concat(checkCompleteness(index))
    .concat(checkUniqueness(index))
    .concat(checkReferences(index))
    .concat(checkDynamicReferences(index))
    .concat(checkFormats(index))
    .concat(checkPickLists(index));

  findings.sort(compareFindings);

  let errorCount = 0;
  let warningCount = 0;
  for (let i = 0; i < findings.length; i++) {
    if (findings[i].severity === 'error') { errorCount++; }
    else if (findings[i].severity === 'warning') { warningCount++; }
  }

  const limit = options.maxFindings === undefined ? DEFAULT_MAX_FINDINGS : options.maxFindings;
  const kept = findings.length > limit ? findings.slice(0, limit) : findings;

  const sheets = summariseSheets(index, findings);

  return {
    fileName: workbook.fileName,
    checkedOn: new Date().toISOString(),
    checkedBy: options.checkedBy,
    findings: kept,
    sheets,
    errorCount,
    warningCount,
    completeness: overallCompleteness(index),
    passed: errorCount === 0
  };
}

/**
 * Parser trouble becomes a finding rather than a console line. A file whose
 * Component sheet failed to decode would otherwise score zero on Component with
 * no indication that the checker never read it.
 */
function readWarningFindings(workbook: ParsedWorkbook): Finding[] {
  return workbook.readWarnings.map((message) => ({
    ruleId: 'structure.unreadable',
    category: 'structure' as const,
    severity: 'error' as const,
    sheet: workbook.fileName,
    message
  }));
}

/**
 * Required-cell completeness for one sheet: the share of cells in required
 * columns that hold a real value.
 *
 * A column the sheet does not have counts as wholly missing rather than being
 * skipped. Skipping it would score a file *higher* for omitting a column than
 * for including it and leaving it blank, which is exactly backwards.
 */
function sheetCompleteness(index: WorkbookIndex, name: CobieSheetName): number | undefined {
  const sheet = index.sheets.get(name);
  if (!sheet || sheet.parsed.rows.length === 0) { return undefined; }

  let required = 0;
  let filled = 0;

  for (let b = 0; b < sheet.bindings.length; b++) {
    const binding = sheet.bindings[b];
    if (binding.column.requirement !== 'required') { continue; }

    required += sheet.parsed.rows.length;
    if (binding.header === undefined) { continue; }

    for (let r = 0; r < sheet.parsed.rows.length; r++) {
      if (!isEmptyOrPlaceholder(sheet.parsed.rows[r].cells[binding.header])) { filled++; }
    }
  }

  if (required === 0) { return undefined; }
  return filled / required;
}

function summariseSheets(index: WorkbookIndex, findings: readonly Finding[]): SheetSummary[] {
  const errors = new Map<string, number>();
  const warnings = new Map<string, number>();

  for (let i = 0; i < findings.length; i++) {
    const finding = findings[i];
    const bucket = finding.severity === 'error' ? errors : finding.severity === 'warning' ? warnings : undefined;
    if (!bucket) { continue; }
    bucket.set(finding.sheet, (bucket.get(finding.sheet) || 0) + 1);
  }

  // Every schema sheet gets a row, present or not, so the dashboard can show an
  // absent required sheet rather than quietly omitting it.
  const names = Object.keys(COBIE_SCHEMA) as CobieSheetName[];
  return names.map((name) => {
    const sheet = index.sheets.get(name);
    return {
      sheet: name,
      present: sheet !== undefined,
      rowCount: sheet ? sheet.parsed.rows.length : 0,
      errors: errors.get(name) || 0,
      warnings: warnings.get(name) || 0,
      completeness: sheetCompleteness(index, name)
    };
  });
}

/**
 * Completeness across the file, weighted by cell rather than by sheet.
 *
 * Per-cell weighting is deliberate: averaging the per-sheet percentages would
 * let a complete four-row Facility sheet offset a threadbare 40,000-row
 * Component sheet, and the headline number would flatter exactly the files that
 * most need work.
 */
function overallCompleteness(index: WorkbookIndex): number {
  let required = 0;
  let filled = 0;

  index.sheets.forEach((sheet) => {
    if (sheet.parsed.rows.length === 0) { return; }

    for (let b = 0; b < sheet.bindings.length; b++) {
      const binding = sheet.bindings[b];
      if (binding.column.requirement !== 'required') { continue; }

      required += sheet.parsed.rows.length;
      if (binding.header === undefined) { continue; }

      for (let r = 0; r < sheet.parsed.rows.length; r++) {
        if (!isEmptyOrPlaceholder(sheet.parsed.rows[r].cells[binding.header])) { filled++; }
      }
    }
  });

  // An empty file is 0% complete, not 100%. `0/0` would otherwise report a
  // perfect score for a workbook with nothing in it.
  if (required === 0) { return 0; }
  return filled / required;
}
