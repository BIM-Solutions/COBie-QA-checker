import type { RawSheet } from '../parse/workbookReader';
import { importReport, isCheckReport } from './reportImport';

/**
 * Fixture grids rather than a round trip through `buildReport`'s real xlsx
 * writer: `write-excel-file`'s `toBlob()` needs a `Blob.arrayBuffer()`, which
 * jsdom - this project's test environment - does not implement.
 *
 * The shape mirrors `summarySheet`/`findingsSheet` in `reportExport.ts`
 * exactly - label/value rows, a blank row, then the per-sheet table on
 * Summary; the seven-column header then one row per finding on Findings - so
 * a change to that layout that this module cannot read back should break this
 * test, the same way `services/validation/fixtures.ts` couples its grids to
 * `sheetFromGrid`.
 */
function reportSheets(): RawSheet[] {
  return [
    {
      name: 'Summary',
      grid: [
        ['COBie check summary'],
        ['File', 'test cobie.xlsx'],
        ['Checked on', '2026-08-30T08:03:48.574Z'],
        ['Checked by', 'andy.hipwood@bimsolutions.ltd.uk'],
        ['Result', 'Failed'],
        ['Errors', '2'],
        ['Warnings', '1'],
        ['Required-field completeness', '75%'],
        [],
        ['Sheet', 'Present', 'Rows', 'Errors', 'Warnings', 'Completeness'],
        ['Contact', 'Yes', '4', '0', '0', '100%'],
        ['Component', 'No', '0', '1', '0', ''],
        ['Space', 'Yes', '10', '1', '0', '82%']
      ]
    },
    {
      name: 'Findings',
      grid: [
        ['Severity', 'Kind', 'Sheet', 'Row', 'Column', 'Value', 'Issue'],
        ['error', 'Structure', 'Component', '', '', '', 'The Component sheet is missing.'],
        ['error', 'Missing values', 'Space', '2', 'Category', '', 'Category is empty.'],
        ['warning', 'Pick list', 'Type', '3', 'AssetType', 'Portable', '"Portable" is not in the AssetType pick list.']
      ]
    }
  ];
}

describe('isCheckReport', () => {
  it('recognises a report shaped like buildReport writes one', () => {
    expect(isCheckReport(reportSheets())).toBe(true);
  });

  it('does not mistake an ordinary COBie workbook for a report', () => {
    const sheets: RawSheet[] = [
      { name: 'Facility', grid: [['Name', 'CreatedBy'], ['Building A', 'a@b.com']] },
      { name: 'Floor', grid: [['Name'], ['L0']] }
    ];
    expect(isCheckReport(sheets)).toBe(false);
  });

  it('requires both the Summary and Findings sheets', () => {
    const summaryOnly: RawSheet[] = [{ name: 'Summary', grid: [['COBie check summary']] }];
    expect(isCheckReport(summaryOnly)).toBe(false);
  });

  it('requires the Findings header to match exactly', () => {
    const sheets = reportSheets();
    const findings = sheets.filter((s) => s.name === 'Findings')[0];
    findings.grid[0] = ['Severity', 'Sheet', 'Row', 'Column', 'Value', 'Issue'];
    expect(isCheckReport(sheets)).toBe(false);
  });
});

describe('importReport', () => {
  it('reconstructs the run the report was built from', () => {
    const run = importReport(reportSheets());

    expect(run.fileName).toBe('test cobie.xlsx');
    expect(run.checkedOn).toBe('2026-08-30T08:03:48.574Z');
    expect(run.checkedBy).toBe('andy.hipwood@bimsolutions.ltd.uk');
    expect(run.errorCount).toBe(2);
    expect(run.warningCount).toBe(1);
    expect(run.completeness).toBeCloseTo(0.75, 2);
    expect(run.passed).toBe(false);

    expect(run.findings).toEqual([
      { ruleId: 'report.imported', category: 'structure', severity: 'error', sheet: 'Component', message: 'The Component sheet is missing.' },
      { ruleId: 'report.imported', category: 'completeness', severity: 'error', sheet: 'Space', row: 2, column: 'Category', message: 'Category is empty.' },
      { ruleId: 'report.imported', category: 'pickList', severity: 'warning', sheet: 'Type', row: 3, column: 'AssetType', value: 'Portable', message: '"Portable" is not in the AssetType pick list.' }
    ]);

    expect(run.sheets).toEqual([
      { sheet: 'Contact', present: true, rowCount: 4, errors: 0, warnings: 0, completeness: 1 },
      { sheet: 'Component', present: false, rowCount: 0, errors: 1, warnings: 0, completeness: undefined },
      { sheet: 'Space', present: true, rowCount: 10, errors: 1, warnings: 0, completeness: 0.82 }
    ]);
  });

  it('treats a sheet-level finding\'s blank row as absent, not zero', () => {
    const run = importReport(reportSheets());
    expect(run.findings[0].row).toBeUndefined();
  });
});
