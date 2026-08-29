import * as React from 'react';
import {
  DetailsList, DetailsListLayoutMode, SelectionMode, Stack, Text
} from '@fluentui/react';
import type { IColumn } from '@fluentui/react';
import styles from '../../app/App.module.scss';
import type { CheckRun, FindingCategory, SheetSummary } from '../../models/findings';
import { CATEGORY_LABELS } from '../../models/findings';
import { useCheckerStore } from '../../state/checkerStore';

/**
 * The roll-up: headline counts, defects by category, and a row per COBie sheet.
 *
 * The category tiles and the sheet rows are filters, not decoration. The
 * question a dashboard raises is always "which of those, and where" and making
 * the answer a click away is most of the tool's value over a static report.
 */

export interface DashboardProps {
  readonly run: CheckRun;
  readonly onDrillDown: () => void;
}

const CATEGORY_ORDER: readonly FindingCategory[] = [
  'structure', 'completeness', 'placeholder', 'reference', 'uniqueness', 'format', 'pickList'
];

interface TileProps {
  readonly label: string;
  readonly value: string;
  readonly tone?: 'error' | 'warning' | 'good';
  readonly onClick?: () => void;
  readonly selected?: boolean;
}

const Tile: React.FC<TileProps> = ({ label, value, tone, onClick, selected }) => {
  const toneClass = tone === 'error' ? styles.error : tone === 'warning' ? styles.warning : tone === 'good' ? styles.good : '';
  const className = [
    styles.tile,
    onClick ? styles.clickable : '',
    selected ? styles.selected : ''
  ].filter((c) => c !== '').join(' ');

  const body = (
    <>
      <div className={`${styles.tileValue} ${toneClass}`}>{value}</div>
      <div className={styles.tileLabel}>{label}</div>
    </>
  );

  // A real <button> when it does something, a <div> when it does not - so a
  // keyboard reaches exactly the tiles that are actionable and no others.
  return onClick
    ? <button type="button" className={className} onClick={onClick}>{body}</button>
    : <div className={className}>{body}</div>;
};

export const Dashboard: React.FC<DashboardProps> = ({ run, onDrillDown }) => {
  const filters = useCheckerStore((s) => s.filters);
  const setFilters = useCheckerStore((s) => s.setFilters);
  const clearFilters = useCheckerStore((s) => s.clearFilters);

  const byCategory = React.useMemo(() => {
    const counts = new Map<FindingCategory, number>();
    for (let i = 0; i < run.findings.length; i++) {
      const finding = run.findings[i];
      if (finding.severity === 'info') { continue; }
      counts.set(finding.category, (counts.get(finding.category) || 0) + 1);
    }
    return counts;
  }, [run]);

  const selectCategory = React.useCallback((category: FindingCategory): void => {
    clearFilters();
    setFilters({ categories: [category] });
    onDrillDown();
  }, [clearFilters, onDrillDown, setFilters]);

  const selectSheet = React.useCallback((sheet: string): void => {
    clearFilters();
    setFilters({ sheets: [sheet] });
    onDrillDown();
  }, [clearFilters, onDrillDown, setFilters]);

  const sheetColumns: IColumn[] = [
    {
      key: 'sheet', name: 'Sheet', minWidth: 110,
      onRender: (item: SheetSummary) => (
        <button
          type="button"
          className={styles.clickable}
          style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer' }}
          onClick={() => selectSheet(item.sheet)}
        >
          {item.sheet}
        </button>
      )
    },
    {
      key: 'present', name: 'Present', minWidth: 70, maxWidth: 80,
      onRender: (item: SheetSummary) => item.present
        ? 'Yes'
        : <span className={styles.subtle}>No</span>
    },
    {
      key: 'rows', name: 'Rows', minWidth: 60, maxWidth: 80, className: styles.rowNumber,
      onRender: (item: SheetSummary) => String(item.rowCount)
    },
    {
      key: 'errors', name: 'Errors', minWidth: 60, maxWidth: 80,
      onRender: (item: SheetSummary) => item.errors > 0
        ? <span className={styles.error}>{item.errors}</span>
        : <span className={styles.subtle}>0</span>
    },
    {
      key: 'warnings', name: 'Warnings', minWidth: 70, maxWidth: 90,
      onRender: (item: SheetSummary) => item.warnings > 0
        ? <span className={styles.warning}>{item.warnings}</span>
        : <span className={styles.subtle}>0</span>
    },
    {
      key: 'completeness', name: 'Complete', minWidth: 80, maxWidth: 100,
      // An em dash, not 0%, when the sheet is absent or empty. A percentage
      // implies a measurement that was never taken.
      onRender: (item: SheetSummary) => item.completeness === undefined
        ? <span className={styles.subtle}>&mdash;</span>
        : `${Math.round(item.completeness * 100)}%`
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <div className={styles.tiles}>
        <Tile
          label={run.passed ? 'No errors found' : 'Errors'}
          value={run.passed ? 'Passed' : String(run.errorCount)}
          tone={run.passed ? 'good' : 'error'}
          onClick={run.errorCount > 0 ? () => { clearFilters(); setFilters({ severities: ['error'] }); onDrillDown(); } : undefined}
          selected={filters.severities.length === 1 && filters.severities[0] === 'error'}
        />
        <Tile
          label="Warnings"
          value={String(run.warningCount)}
          tone={run.warningCount > 0 ? 'warning' : undefined}
          onClick={run.warningCount > 0 ? () => { clearFilters(); setFilters({ severities: ['warning'] }); onDrillDown(); } : undefined}
          selected={filters.severities.length === 1 && filters.severities[0] === 'warning'}
        />
        <Tile
          label="Required fields complete"
          value={`${Math.round(run.completeness * 100)}%`}
          tone={run.completeness >= 0.95 ? 'good' : run.completeness >= 0.8 ? 'warning' : 'error'}
        />
        <Tile
          label="Sheets present"
          value={`${run.sheets.filter((s) => s.present).length} of ${run.sheets.length}`}
        />
      </div>

      <div className={styles.section}>Defects by kind</div>
      <div className={styles.tiles}>
        {CATEGORY_ORDER.filter((category) => (byCategory.get(category) || 0) > 0).map((category) => (
          <Tile
            key={category}
            label={CATEGORY_LABELS[category]}
            value={String(byCategory.get(category) || 0)}
            onClick={() => selectCategory(category)}
            selected={filters.categories.length === 1 && filters.categories[0] === category}
          />
        ))}
        {byCategory.size === 0 && (
          <Text className={styles.subtle}>Nothing to report beyond the notes below.</Text>
        )}
      </div>

      <div className={styles.section}>By sheet</div>
      <div className={styles.tableScroll}>
        <DetailsList
          items={run.sheets as SheetSummary[]}
          columns={sheetColumns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          compact
        />
      </div>
    </Stack>
  );
};
