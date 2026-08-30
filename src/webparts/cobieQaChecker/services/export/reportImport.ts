import type { CheckRun, Finding, FindingCategory, Severity, SheetSummary } from '../../models/findings';
import { CATEGORY_LABELS } from '../../models/findings';
import type { RawSheet } from '../parse';

/**
 * Reading the checker's own exported report back into a `CheckRun`.
 *
 * The report lives beside the source file it describes (`reportFolder`
 * defaults to `sourceFolder`), so it shows up in the same "files in this
 * library" list a person picks a COBie deliverable from. Without this, opening
 * a report there falls into the normal check path: it gets read as a COBie
 * file, every sheet the schema expects is "missing", and `Summary`/`Findings`
 * are reported as sheets the checker does not recognise - the previous run's
 * results, nowhere to be seen.
 *
 * Detection is by content, not file name: a report renamed by whoever exported
 * it must still be recognised, and a genuine COBie file must never be
 * misread as one.
 */

const FINDINGS_HEADER: readonly string[] = ['Severity', 'Kind', 'Sheet', 'Row', 'Column', 'Value', 'Issue'];
const SEVERITIES: readonly Severity[] = ['error', 'warning', 'info'];

const CATEGORY_BY_LABEL: Readonly<Record<string, FindingCategory>> = (() => {
  const map: Record<string, FindingCategory> = {};
  (Object.keys(CATEGORY_LABELS) as FindingCategory[]).forEach((category) => {
    map[CATEGORY_LABELS[category]] = category;
  });
  return map;
})();

function cell(row: string[] | undefined, index: number): string {
  const value = row ? row[index] : undefined;
  return value === undefined || value === null ? '' : String(value).trim();
}

function findSheet(sheets: readonly RawSheet[], name: string): RawSheet | undefined {
  return sheets.filter((sheet) => sheet.name === name)[0];
}

/** True when `sheets` is shaped exactly as `buildReport` writes it. */
export function isCheckReport(sheets: readonly RawSheet[]): boolean {
  const summary = findSheet(sheets, 'Summary');
  const findings = findSheet(sheets, 'Findings');
  if (!summary || !findings) { return false; }

  if (cell(summary.grid[0], 0) !== 'COBie check summary') { return false; }

  const header = findings.grid[0] || [];
  if (header.length !== FINDINGS_HEADER.length) { return false; }
  for (let i = 0; i < FINDINGS_HEADER.length; i++) {
    if (cell(header, i) !== FINDINGS_HEADER[i]) { return false; }
  }
  return true;
}

/**
 * The label/value rows above the per-sheet table: `File`, `Checked on`, and so
 * on. Read by label rather than by fixed row index so an unrelated reordering
 * of `summarySheet` cannot silently misattribute a value.
 */
function readFields(grid: string[][]): Map<string, string> {
  const fields = new Map<string, string>();
  for (let i = 1; i < grid.length; i++) {
    const label = cell(grid[i], 0);
    if (label === '') { break; }
    fields.set(label, cell(grid[i], 1));
  }
  return fields;
}

function readSheetSummaries(grid: string[][]): SheetSummary[] {
  let headerRow = -1;
  for (let i = 0; i < grid.length; i++) {
    if (cell(grid[i], 0) === 'Sheet' && cell(grid[i], 1) === 'Present') { headerRow = i; break; }
  }
  if (headerRow === -1) { return []; }

  const summaries: SheetSummary[] = [];
  for (let i = headerRow + 1; i < grid.length; i++) {
    const name = cell(grid[i], 0);
    if (name === '') { continue; }
    const completenessText = cell(grid[i], 5).replace('%', '');
    summaries.push({
      sheet: name,
      present: cell(grid[i], 1) === 'Yes',
      rowCount: Number(cell(grid[i], 2)) || 0,
      errors: Number(cell(grid[i], 3)) || 0,
      warnings: Number(cell(grid[i], 4)) || 0,
      completeness: completenessText === '' ? undefined : Number(completenessText) / 100
    });
  }
  return summaries;
}

function readFindings(grid: string[][]): Finding[] {
  const findings: Finding[] = [];
  for (let i = 1; i < grid.length; i++) {
    const severity = cell(grid[i], 0).toLowerCase() as Severity;
    if (SEVERITIES.indexOf(severity) === -1) { continue; }

    const row = cell(grid[i], 3);
    const column = cell(grid[i], 4);
    const value = cell(grid[i], 5);

    findings.push({
      // Not the original rule id - the report never wrote it, deliberately
      // (see `models/findings.ts`: "Not shown to users"). Nothing here reads
      // it back, so a fixed placeholder costs nothing.
      ruleId: 'report.imported',
      category: CATEGORY_BY_LABEL[cell(grid[i], 1)] || 'structure',
      severity,
      sheet: cell(grid[i], 2),
      row: row === '' ? undefined : Number(row),
      column: column === '' ? undefined : column,
      value: value === '' ? undefined : value,
      message: cell(grid[i], 6)
    });
  }
  return findings;
}

/** Reconstructs the `CheckRun` a previously exported report was built from. */
export function importReport(sheets: readonly RawSheet[]): CheckRun {
  const summary = findSheet(sheets, 'Summary');
  const findingsSheet = findSheet(sheets, 'Findings');
  if (!summary || !findingsSheet) {
    throw new Error('Not a COBie check report.');
  }

  const fields = readFields(summary.grid);
  const completenessText = (fields.get('Required-field completeness') || '').replace('%', '');

  return {
    fileName: fields.get('File') || '',
    checkedOn: fields.get('Checked on') || '',
    checkedBy: fields.get('Checked by') || '',
    findings: readFindings(findingsSheet.grid),
    sheets: readSheetSummaries(summary.grid),
    errorCount: Number(fields.get('Errors')) || 0,
    warningCount: Number(fields.get('Warnings')) || 0,
    completeness: completenessText === '' ? 0 : Number(completenessText) / 100,
    passed: fields.get('Result') === 'Passed'
  };
}
