import * as React from 'react';
import {
  DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, SelectionMode,
  SearchBox, Stack, Text
} from '@fluentui/react';
import type { IColumn, IDropdownOption } from '@fluentui/react';
import styles from '../../app/App.module.scss';
import type { CheckRun, Finding, FindingCategory, Severity } from '../../models/findings';
import { CATEGORY_LABELS } from '../../models/findings';
import { applyFilters, useCheckerStore } from '../../state/checkerStore';

/**
 * The work list.
 *
 * Virtualised by DetailsList, which matters: the engine caps findings at 5,000
 * and rendering that many rows unvirtualised locks the tab. Rows are rendered
 * as plain text rather than links because there is nowhere to link *to* - the
 * source is a spreadsheet in a library, and the row number is what a person
 * takes back to Excel.
 */

export interface IssuesTableProps {
  readonly run: CheckRun;
}

const SEVERITIES: readonly Severity[] = ['error', 'warning', 'info'];

const CATEGORIES: readonly FindingCategory[] = [
  'structure', 'completeness', 'placeholder', 'reference', 'uniqueness', 'format', 'pickList'
];

export const IssuesTable: React.FC<IssuesTableProps> = ({ run }) => {
  const filters = useCheckerStore((s) => s.filters);
  const setFilters = useCheckerStore((s) => s.setFilters);
  const clearFilters = useCheckerStore((s) => s.clearFilters);

  const visible = React.useMemo(
    () => applyFilters(run.findings, filters),
    [run.findings, filters]
  );

  // Only sheets that actually have findings: offering all eighteen when four
  // have defects makes the user hunt for the ones worth picking.
  const sheetOptions: IDropdownOption[] = React.useMemo(() => {
    const names: string[] = [];
    for (let i = 0; i < run.findings.length; i++) {
      if (names.indexOf(run.findings[i].sheet) === -1) { names.push(run.findings[i].sheet); }
    }
    names.sort();
    return names.map((name) => ({ key: name, text: name }));
  }, [run.findings]);

  const columns: IColumn[] = [
    {
      key: 'severity', name: 'Severity', minWidth: 70, maxWidth: 90,
      onRender: (item: Finding) => (
        <span className={item.severity === 'error' ? styles.error : item.severity === 'warning' ? styles.warning : styles.subtle}>
          {item.severity === 'error' ? 'Error' : item.severity === 'warning' ? 'Warning' : 'Note'}
        </span>
      )
    },
    { key: 'sheet', name: 'Sheet', fieldName: 'sheet', minWidth: 90, maxWidth: 120 },
    {
      key: 'row', name: 'Row', minWidth: 50, maxWidth: 60, className: styles.rowNumber,
      // Blank for a sheet-level finding. Rendering 0 would send someone looking
      // for a row that does not exist.
      onRender: (item: Finding) => item.row === undefined
        ? <span className={styles.subtle}>&mdash;</span>
        : String(item.row)
    },
    {
      key: 'column', name: 'Column', minWidth: 100, maxWidth: 150,
      onRender: (item: Finding) => item.column || ''
    },
    {
      key: 'category', name: 'Kind', minWidth: 100, maxWidth: 130,
      onRender: (item: Finding) => CATEGORY_LABELS[item.category]
    },
    {
      key: 'message', name: 'Issue', fieldName: 'message', minWidth: 240,
      isMultiline: true, isResizable: true
    }
  ];

  const filtered = filters.severities.length > 0 || filters.categories.length > 0 ||
    filters.sheets.length > 0 || filters.search !== '';

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <div className={styles.filters}>
        <Dropdown
          label="Severity"
          multiSelect
          selectedKeys={filters.severities as string[]}
          options={SEVERITIES.map((s) => ({
            key: s,
            text: s === 'error' ? 'Error' : s === 'warning' ? 'Warning' : 'Note'
          }))}
          styles={{ root: { minWidth: 150 } }}
          onChange={(_, option) => {
            if (!option) { return; }
            const key = option.key as Severity;
            const next = option.selected
              ? filters.severities.concat([key])
              : filters.severities.filter((s) => s !== key);
            setFilters({ severities: next });
          }}
        />
        <Dropdown
          label="Kind"
          multiSelect
          selectedKeys={filters.categories as string[]}
          options={CATEGORIES.map((c) => ({ key: c, text: CATEGORY_LABELS[c] }))}
          styles={{ root: { minWidth: 190 } }}
          onChange={(_, option) => {
            if (!option) { return; }
            const key = option.key as FindingCategory;
            const next = option.selected
              ? filters.categories.concat([key])
              : filters.categories.filter((c) => c !== key);
            setFilters({ categories: next });
          }}
        />
        <Dropdown
          label="Sheet"
          multiSelect
          selectedKeys={filters.sheets as string[]}
          options={sheetOptions}
          styles={{ root: { minWidth: 170 } }}
          onChange={(_, option) => {
            if (!option) { return; }
            const key = option.key as string;
            const next = option.selected
              ? filters.sheets.concat([key])
              : filters.sheets.filter((s) => s !== key);
            setFilters({ sheets: next });
          }}
        />
        <SearchBox
          placeholder="Search values and messages"
          value={filters.search}
          styles={{ root: { minWidth: 220 } }}
          onChange={(_, value) => setFilters({ search: value || '' })}
        />
        {filtered && (
          <DefaultButton text="Clear filters" iconProps={{ iconName: 'ClearFilter' }} onClick={clearFilters} />
        )}
      </div>

      <Text className={styles.subtle}>
        {visible.length === run.findings.length
          ? `${run.findings.length} finding${run.findings.length === 1 ? '' : 's'}`
          : `${visible.length} of ${run.findings.length} findings`}
        {run.findings.length < run.errorCount + run.warningCount &&
          ` — the list is capped; the counts above are the true totals.`}
      </Text>

      {visible.length === 0
        ? (
          <div className={styles.empty}>
            {run.findings.length === 0
              ? 'No findings. This file passes every check.'
              : 'No findings match these filters.'}
          </div>
        )
        : (
          <div className={styles.tableScroll}>
            <DetailsList
              items={visible}
              columns={columns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              compact
            />
          </div>
        )}
    </Stack>
  );
};
