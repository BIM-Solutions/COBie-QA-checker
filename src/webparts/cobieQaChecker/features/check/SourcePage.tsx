import * as React from 'react';
import {
  DefaultButton, DetailsList, DetailsListLayoutMode, SelectionMode,
  MessageBar, MessageBarType, Spinner, SpinnerSize, Stack, Text
} from '@fluentui/react';
import type { IColumn } from '@fluentui/react';
import styles from '../../app/App.module.scss';
import type { CheckerServices } from '../../services/CheckerServices';
import type { LibraryFile, RunHistoryEntry } from '../../services/sharepoint';
import { useCheckerStore } from '../../state/checkerStore';
import type { UseCheck } from '../../app/useCheck';

/**
 * Choosing what to check.
 *
 * Two routes, and both matter. The library list is the point of the web part -
 * the deliverable is already in SharePoint and nobody should have to download
 * it. Local upload stays because the most common moment to want a COBie check
 * is *before* the file is issued, when it is still on someone's machine.
 */

export interface SourcePageProps {
  readonly services: CheckerServices;
  readonly check: UseCheck;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) { return `${bytes} B`; }
  if (bytes < 1024 * 1024) { return `${Math.round(bytes / 1024)} KB`; }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso: string): string {
  if (!iso) { return ''; }
  const date = new Date(iso);
  return isNaN(date.getTime()) ? iso : date.toLocaleString();
}

export const SourcePage: React.FC<SourcePageProps> = ({ services, check }) => {
  const busy = useCheckerStore((s) => s.busy);
  const busyMessage = useCheckerStore((s) => s.busyMessage);

  const [files, setFiles] = React.useState<LibraryFile[]>([]);
  const [history, setHistory] = React.useState<RunHistoryEntry[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [listError, setListError] = React.useState<string | undefined>();
  const [dragging, setDragging] = React.useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    let cancelled = false;

    async function load(): Promise<void> {
      setLoading(true);
      setListError(undefined);
      try {
        const found = await services.files.listFiles(services.sourceFolder);
        if (!cancelled) { setFiles(found); }
      } catch (error) {
        // A misconfigured library is the most likely first-run problem, so the
        // message names the folder rather than saying "something went wrong".
        if (!cancelled) {
          setListError(error instanceof Error ? error.message : String(error));
        }
      } finally {
        if (!cancelled) { setLoading(false); }
      }

      if (!services.historyEnabled) { return; }
      try {
        const recent = await services.history.recentRuns(5);
        if (!cancelled) { setHistory(recent); }
      } catch {
        // History is a nicety. A site where the list cannot be read still
        // checks files perfectly well, so this failure is swallowed rather
        // than shown.
        if (!cancelled) { setHistory([]); }
      }
    }

    void load();
    return () => { cancelled = true; };
  }, [services]);

  const onDrop = React.useCallback((event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    setDragging(false);
    const dropped = event.dataTransfer.files;
    if (dropped && dropped.length > 0) { void check.checkLocalFile(dropped[0]); }
  }, [check]);

  const onPick = React.useCallback((event: React.ChangeEvent<HTMLInputElement>): void => {
    const chosen = event.target.files;
    if (chosen && chosen.length > 0) { void check.checkLocalFile(chosen[0]); }
    // Cleared so picking the same file twice in a row fires a second change
    // event; without this, re-checking a file after editing it does nothing.
    event.target.value = '';
  }, [check]);

  const fileColumns: IColumn[] = [
    {
      key: 'name', name: 'File', fieldName: 'name', minWidth: 180, isMultiline: true,
      onRender: (item: LibraryFile) => (
        <DefaultButton
          text={item.name}
          styles={{ root: { border: 'none', background: 'none', padding: 0, height: 'auto', textAlign: 'left' } }}
          disabled={busy}
          onClick={() => { void check.checkLibraryFile(item); }}
        />
      )
    },
    { key: 'size', name: 'Size', minWidth: 70, maxWidth: 90, onRender: (i: LibraryFile) => formatSize(i.size) },
    { key: 'modified', name: 'Modified', minWidth: 140, onRender: (i: LibraryFile) => formatDate(i.modified) },
    { key: 'by', name: 'By', minWidth: 120, onRender: (i: LibraryFile) => i.modifiedBy || '' }
  ];

  const historyColumns: IColumn[] = [
    { key: 'file', name: 'File', fieldName: 'fileName', minWidth: 160, isMultiline: true },
    { key: 'when', name: 'Checked', minWidth: 140, onRender: (i: RunHistoryEntry) => formatDate(i.checkedOn) },
    {
      key: 'result', name: 'Result', minWidth: 80,
      onRender: (i: RunHistoryEntry) => (
        <span className={i.passed ? styles.good : styles.error}>{i.passed ? 'Passed' : 'Failed'}</span>
      )
    },
    { key: 'errors', name: 'Errors', minWidth: 60, onRender: (i: RunHistoryEntry) => String(i.errors) },
    {
      key: 'complete', name: 'Complete', minWidth: 80,
      onRender: (i: RunHistoryEntry) => `${Math.round(i.completeness * 100)}%`
    }
  ];

  if (busy) {
    return (
      <Stack tokens={{ childrenGap: 12 }} horizontalAlign="center" style={{ padding: 40 }}>
        <Spinner size={SpinnerSize.large} label={busyMessage || 'Working'} />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <div
        className={`${styles.dropZone} ${dragging ? styles.dropZoneActive : ''}`}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
      >
        <Text block>Drop a COBie file here to check it without uploading it</Text>
        <Text block className={styles.subtle}>.xlsx, .xlsm or .csv</Text>
        <div style={{ marginTop: 12 }}>
          <DefaultButton
            text="Choose a file"
            iconProps={{ iconName: 'OpenFile' }}
            onClick={() => inputRef.current?.click()}
          />
        </div>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xlsm,.xls,.csv,.tsv,.txt"
          style={{ display: 'none' }}
          onChange={onPick}
        />
      </div>

      <div>
        <div className={styles.section}>Files in {services.sourceFolder}</div>

        {listError !== undefined && (
          <MessageBar messageBarType={MessageBarType.error}>{listError}</MessageBar>
        )}

        {loading && <Spinner label="Loading files" />}

        {!loading && listError === undefined && files.length === 0 && (
          <div className={styles.empty}>
            No spreadsheets in this library. Set a different one in the web part properties,
            or drop a file above.
          </div>
        )}

        {!loading && files.length > 0 && (
          <div className={styles.tableScroll}>
            <DetailsList
              items={files}
              columns={fileColumns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              compact
            />
          </div>
        )}
      </div>

      {history.length > 0 && (
        <div>
          <div className={styles.section}>Recent checks on this site</div>
          <div className={styles.tableScroll}>
            <DetailsList
              items={history}
              columns={historyColumns}
              selectionMode={SelectionMode.none}
              layoutMode={DetailsListLayoutMode.justified}
              compact
            />
          </div>
        </div>
      )}
    </Stack>
  );
};
