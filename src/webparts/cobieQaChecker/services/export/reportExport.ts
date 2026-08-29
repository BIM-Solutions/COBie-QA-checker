import type { CheckRun, Finding } from '../../models/findings';
import { CATEGORY_LABELS } from '../../models/findings';

/**
 * The findings report, as an Excel workbook.
 *
 * Excel rather than PDF because the output is a work list: someone filters it
 * by sheet, assigns rows, and ticks them off. A PDF would look more finished
 * and be less useful.
 *
 * The writer is a dynamic import for the same reason the reader is - it is a
 * large dependency that a page merely hosting the checker should not pay for.
 */

/**
 * The writer's own cell type is generic over four value shapes and expresses
 * "a bare value or an object" as a union, which makes every row literal here
 * need a cast. This is the same shape, narrowed to what the report uses.
 */
interface Cell {
  value?: string | number;
  fontWeight?: 'bold';
  backgroundColor?: string;
  wrap?: boolean;
  align?: 'left' | 'center' | 'right';
}

const HEADER: Cell = { fontWeight: 'bold', backgroundColor: '#f3f2f1' };

function headerRow(labels: readonly string[]): Cell[] {
  return labels.map((value) => ({ ...HEADER, value }));
}

/**
 * Two sheets: a summary a manager reads and a findings list an engineer works.
 * Splitting them means neither audience has to scroll past the other's content.
 */
function summarySheet(run: CheckRun): Cell[][] {
  const rows: Cell[][] = [
    [{ value: 'COBie check summary', fontWeight: 'bold' }],
    [{ value: 'File' }, { value: run.fileName }],
    [{ value: 'Checked on' }, { value: run.checkedOn }],
    [{ value: 'Checked by' }, { value: run.checkedBy }],
    [{ value: 'Result' }, { value: run.passed ? 'Passed' : 'Failed' }],
    [{ value: 'Errors' }, { value: run.errorCount }],
    [{ value: 'Warnings' }, { value: run.warningCount }],
    [{ value: 'Required-field completeness' }, { value: `${Math.round(run.completeness * 100)}%` }],
    [],
    headerRow(['Sheet', 'Present', 'Rows', 'Errors', 'Warnings', 'Completeness'])
  ];

  for (let i = 0; i < run.sheets.length; i++) {
    const sheet = run.sheets[i];
    rows.push([
      { value: sheet.sheet },
      { value: sheet.present ? 'Yes' : 'No' },
      { value: sheet.rowCount },
      { value: sheet.errors },
      { value: sheet.warnings },
      // An absent sheet gets an empty cell, not 0%. Writing 0% would assert a
      // measurement that was never taken.
      { value: sheet.completeness === undefined ? '' : `${Math.round(sheet.completeness * 100)}%` }
    ]);
  }

  return rows;
}

function findingsSheet(findings: readonly Finding[]): Cell[][] {
  const rows: Cell[][] = [
    headerRow(['Severity', 'Kind', 'Sheet', 'Row', 'Column', 'Value', 'Issue'])
  ];

  for (let i = 0; i < findings.length; i++) {
    const finding = findings[i];
    rows.push([
      { value: finding.severity },
      { value: CATEGORY_LABELS[finding.category] },
      { value: finding.sheet },
      // Blank rather than 0 for a sheet-level finding: row 0 does not exist and
      // would send someone looking for it.
      { value: finding.row === undefined ? '' : finding.row },
      { value: finding.column || '' },
      { value: finding.value || '' },
      { value: finding.message, wrap: true }
    ]);
  }

  return rows;
}

export interface ReportBlob {
  readonly fileName: string;
  readonly buffer: ArrayBuffer;
}

/** A file name that sorts by date and cannot collide across two runs a minute apart. */
export function reportFileName(run: CheckRun): string {
  const base = run.fileName.replace(/\.[^.]+$/, '');
  const stamp = run.checkedOn.replace(/[:.]/g, '-').replace(/Z$/, '');
  return `${base} - COBie check ${stamp}.xlsx`;
}

export async function buildReport(run: CheckRun): Promise<ReportBlob> {
  // `/browser` rather than the bare package: the root exposes no CommonJS entry,
  // and the browser build is the one that produces a Blob without Node's
  // streams. The dynamic import keeps it out of the entry bundle.
  const writeXlsxFile = (await import(
    /* webpackChunkName: "xlsx-writer" */ 'write-excel-file/browser'
  )).default;

  // Widths chosen so the message column is readable without the reviewer
  // widening it by hand on every report.
  const blob = await writeXlsxFile([
    {
      sheet: 'Summary',
      data: summarySheet(run),
      columns: [{ width: 32 }, { width: 60 }]
    },
    {
      sheet: 'Findings',
      data: findingsSheet(run.findings),
      columns: [
        { width: 10 }, { width: 18 }, { width: 14 }, { width: 8 },
        { width: 22 }, { width: 28 }, { width: 90 }
      ]
    }
  ]).toBlob();

  return { fileName: reportFileName(run), buffer: await blob.arrayBuffer() };
}
