import type { CobieSheetName } from './cobieSchema';
import type { SourceRow } from './workbook';

/**
 * What the checker reports, and how a run rolls up.
 *
 * A finding is always anchored to somewhere a person can go and look: a sheet,
 * usually a row, usually a column. Rules that cannot say where a problem is —
 * "the Space sheet is missing" — anchor to the sheet alone, and the UI renders
 * that difference rather than showing a blank row number.
 */

export type Severity = 'error' | 'warning' | 'info';

/**
 * Rule families, used for grouping in the UI and for the per-category roll-up.
 * Kept coarse on purpose: a report that sorts 4,000 findings into 30 buckets is
 * no more actionable than one long list.
 */
export type FindingCategory =
  | 'structure'
  | 'completeness'
  | 'placeholder'
  | 'uniqueness'
  | 'reference'
  | 'format'
  | 'pickList';

export interface Finding {
  /** Stable rule identifier, e.g. `reference.broken`. Not shown to users. */
  readonly ruleId: string;
  readonly category: FindingCategory;
  readonly severity: Severity;
  readonly sheet: string;
  /** Absent for sheet-level findings. */
  readonly row?: SourceRow;
  /** Absent for sheet- and row-level findings. */
  readonly column?: string;
  /** The offending value, when quoting it helps. Never invented. */
  readonly value?: string;
  /** One sentence, sentence case, naming what is wrong. */
  readonly message: string;
}

/** Per-sheet roll-up. `expected` sheets absent from the file still get an entry. */
export interface SheetSummary {
  readonly sheet: string;
  readonly present: boolean;
  readonly rowCount: number;
  readonly errors: number;
  readonly warnings: number;
  /**
   * Share of required cells on this sheet that hold a real value: not blank,
   * not a placeholder. Undefined when the sheet is absent or has no rows, where
   * a percentage would imply a measurement that was never taken.
   */
  readonly completeness?: number;
}

export interface CheckRun {
  readonly fileName: string;
  /** ISO 8601, UTC. */
  readonly checkedOn: string;
  readonly checkedBy: string;
  readonly findings: readonly Finding[];
  readonly sheets: readonly SheetSummary[];
  readonly errorCount: number;
  readonly warningCount: number;
  /**
   * Required-cell completeness across the whole file, 0–1. This is the headline
   * number, and it counts only *required* columns on sheets that are present:
   * padding it with optional columns would flatter every file that ships the
   * bare minimum, which is the opposite of what a QC tool is for.
   */
  readonly completeness: number;
  /** False when any error-severity finding exists. */
  readonly passed: boolean;
}

export const SEVERITY_ORDER: Readonly<Record<Severity, number>> = {
  error: 0,
  warning: 1,
  info: 2
};

export const CATEGORY_LABELS: Readonly<Record<FindingCategory, string>> = {
  structure: 'Structure',
  completeness: 'Missing values',
  placeholder: 'Placeholder values',
  uniqueness: 'Duplicates',
  reference: 'Broken references',
  format: 'Format',
  pickList: 'Pick list'
};

/** Errors first, then by sheet, then by row — reading order for a fix list. */
export function compareFindings(a: Finding, b: Finding): number {
  const bySeverity = SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity];
  if (bySeverity !== 0) { return bySeverity; }
  if (a.sheet !== b.sheet) { return a.sheet < b.sheet ? -1 : 1; }
  return (a.row || 0) - (b.row || 0);
}

export function isCobieSheetName(sheet: string, names: readonly string[]): sheet is CobieSheetName {
  return names.indexOf(sheet) !== -1;
}
