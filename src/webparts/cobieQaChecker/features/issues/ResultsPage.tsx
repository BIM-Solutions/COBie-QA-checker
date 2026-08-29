import * as React from 'react';
import { DefaultButton, Pivot, PivotItem, PrimaryButton, Stack, Text } from '@fluentui/react';
import styles from '../../app/App.module.scss';
import type { CheckRun } from '../../models/findings';
import { Dashboard } from '../dashboard/Dashboard';
import { IssuesTable } from './IssuesTable';
import { useCheckerStore } from '../../state/checkerStore';
import type { UseCheck } from '../../app/useCheck';

/**
 * What a finished run looks like: a header saying what was checked and how it
 * did, then the roll-up and the work list as two tabs.
 *
 * Tabs rather than one long page because the two views serve different people
 * at different moments, and stacking them would put a 5,000-row table under
 * every visit to the summary.
 */

export interface ResultsPageProps {
  readonly run: CheckRun;
  readonly check: UseCheck;
}

type Tab = 'summary' | 'issues';

export const ResultsPage: React.FC<ResultsPageProps> = ({ run, check }) => {
  const reset = useCheckerStore((s) => s.reset);
  const busy = useCheckerStore((s) => s.busy);
  const [tab, setTab] = React.useState<Tab>('summary');

  const drillDown = React.useCallback(() => setTab('issues'), []);

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <div className={styles.header}>
        <div className={styles.headline}>
          <Text variant="large" className={styles.fileName}>{run.fileName}</Text>
          <Text className={run.passed ? styles.good : styles.error}>
            {run.passed ? 'Passed' : `Failed — ${run.errorCount} error${run.errorCount === 1 ? '' : 's'}`}
          </Text>
          <Text className={styles.subtle}>
            Checked {new Date(run.checkedOn).toLocaleString()} by {run.checkedBy}
          </Text>
        </div>
        <div className={styles.actions}>
          <PrimaryButton
            text="Export report"
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={busy}
            onClick={() => { void check.exportReport(); }}
          />
          <DefaultButton text="Check another file" iconProps={{ iconName: 'Back' }} onClick={reset} />
        </div>
      </div>

      <Pivot
        selectedKey={tab}
        onLinkClick={(item) => setTab((item?.props.itemKey as Tab) || 'summary')}
      >
        <PivotItem headerText="Summary" itemKey="summary">
          <div style={{ paddingTop: 12 }}>
            <Dashboard run={run} onDrillDown={drillDown} />
          </div>
        </PivotItem>
        <PivotItem
          headerText="Findings"
          itemKey="issues"
          // The badge is the errors-plus-warnings total, not the capped list
          // length: a tab reading "5000" on a file with 90,000 defects would
          // understate the problem at exactly the moment it matters most.
          itemCount={run.errorCount + run.warningCount}
        >
          <div style={{ paddingTop: 12 }}>
            <IssuesTable run={run} />
          </div>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};
