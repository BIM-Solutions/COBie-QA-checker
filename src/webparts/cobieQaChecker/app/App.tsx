import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import styles from './App.module.scss';
import type { CheckerServices } from '../services/CheckerServices';
import { useCheckerStore } from '../state/checkerStore';
import { useCheck } from './useCheck';
import { SourcePage } from '../features/check/SourcePage';
import { ResultsPage } from '../features/issues/ResultsPage';

/**
 * The shell.
 *
 * Two screens and no router: SPFx web parts share the page's URL with whatever
 * else is on it, so pushing routes into the address bar would fight the host
 * page. The store's `screen` is the whole of the navigation model.
 */

export interface AppProps {
  readonly services: CheckerServices;
}

export const App: React.FC<AppProps> = ({ services }) => {
  const screen = useCheckerStore((s) => s.screen);
  const run = useCheckerStore((s) => s.run);
  const error = useCheckerStore((s) => s.error);
  const notice = useCheckerStore((s) => s.notice);
  const setError = useCheckerStore((s) => s.setError);
  const setNotice = useCheckerStore((s) => s.setNotice);

  const check = useCheck(services);

  return (
    <div className={styles.app}>
      <Stack tokens={{ childrenGap: 12 }}>
        {error !== undefined && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setError(undefined)}
          >
            {error}
          </MessageBar>
        )}

        {notice !== undefined && (
          <MessageBar
            messageBarType={MessageBarType.success}
            onDismiss={() => setNotice(undefined)}
          >
            {notice}
          </MessageBar>
        )}

        {/*
          `run` is checked as well as `screen` so a results screen can never
          render without a run behind it - the store allows the combination and
          a crash here would take the whole page down with it.
        */}
        {screen === 'results' && run
          ? <ResultsPage run={run} check={check} />
          : <SourcePage services={services} check={check} />}
      </Stack>
    </div>
  );
};
