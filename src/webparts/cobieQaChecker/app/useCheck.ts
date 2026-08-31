import { useCallback } from 'react';
import type { CheckerServices } from '../services/CheckerServices';
import type { LibraryFile } from '../services/sharepoint';
import { isDelimitedFileName, readWorkbook, readXlsxSheets, sheetFromGrid } from '../services/parse';
import type { FileSource } from '../services/parse';
import { runChecks } from '../services/validation';
import { buildReport, importReport, isCheckReport } from '../services/export';
import type { CheckRun } from '../models/findings';
import type { ParsedWorkbook } from '../models/workbook';
import { useCheckerStore } from '../state/checkerStore';

/**
 * The one place a check actually happens.
 *
 * Everything the screens do funnels through here so that the sequence - read,
 * validate, record, show - exists once. The two check entry points differ only
 * in where the bytes come from.
 *
 * One branch interrupts that sequence: a file that turns out to be the
 * checker's own exported report is restored rather than checked. See
 * `readSource`.
 */

type ReadResult =
  | { readonly kind: 'report'; readonly run: CheckRun }
  | { readonly kind: 'workbook'; readonly parsed: ParsedWorkbook };

/**
 * Recognises the checker's own exported report before treating the file as a
 * COBie deliverable.
 *
 * `reportFolder` defaults to `sourceFolder`, so an exported report sits right
 * next to the file it describes and shows up in the same "files in this
 * library" list. Without this, opening it here fell into the normal check: it
 * was read as a COBie file, every sheet the schema expects was reported
 * missing, and the previous run's own results were nowhere to be seen - which
 * is the bug this guards against. Xlsx sheets are read once, raw, so the file
 * is not parsed twice on the far more common path where it is not a report.
 */
async function readSource(source: FileSource): Promise<ReadResult> {
  if (isDelimitedFileName(source.name)) {
    return { kind: 'workbook', parsed: await readWorkbook(source) };
  }

  try {
    const raw = await readXlsxSheets(await source.arrayBuffer());
    if (isCheckReport(raw)) {
      return { kind: 'report', run: importReport(raw) };
    }
    return {
      kind: 'workbook',
      parsed: {
        fileName: source.name,
        sheets: raw.map((sheet) => sheetFromGrid(sheet.name, sheet.grid)),
        readWarnings: []
      }
    };
  } catch (error) {
    // Matches `readWorkbook`'s own handling of an unreadable xlsx file: a
    // finding explaining why, not a thrown error the user cannot act on.
    return {
      kind: 'workbook',
      parsed: {
        fileName: source.name,
        sheets: [],
        readWarnings: [
          `The file could not be read as a spreadsheet: ` +
          `${error instanceof Error ? error.message : String(error)}`
        ]
      }
    };
  }
}

export interface UseCheck {
  checkLibraryFile(file: LibraryFile): Promise<void>;
  checkLocalFile(file: File): Promise<void>;
  exportReport(): Promise<void>;
}

/** Wraps a browser File so the reader does not know where the bytes came from. */
function fromLocalFile(file: File): FileSource {
  return {
    name: file.name,
    arrayBuffer: () => file.arrayBuffer(),
    text: () => file.text()
  };
}

export function useCheck(services: CheckerServices): UseCheck {
  const setRun = useCheckerStore((s) => s.setRun);
  const setBusy = useCheckerStore((s) => s.setBusy);
  const setError = useCheckerStore((s) => s.setError);
  const setNotice = useCheckerStore((s) => s.setNotice);

  const check = useCallback(async (source: FileSource): Promise<void> => {
    setError(undefined);
    setNotice(undefined);
    setBusy(true, `Reading ${source.name}`);

    try {
      const read = await readSource(source);

      if (read.kind === 'report') {
        setRun(read.run);
        setBusy(false);
        setNotice(
          `This is a COBie check report exported earlier, not a COBie deliverable - ` +
          `showing the results it already recorded rather than checking it again.`
        );
        return;
      }

      setBusy(true, 'Checking against the COBie schema');
      // One yield before the synchronous rule pass, so the browser paints the
      // message above before the main thread is blocked. Without it the user
      // sees "Reading" throughout and the app looks stuck on a large file.
      await new Promise((resolve) => setTimeout(resolve, 0));

      const run = runChecks(read.parsed, { checkedBy: services.currentUser });
      setRun(run);
      setBusy(false);

      if (!services.historyEnabled) { return; }

      try {
        await services.history.recordRun(run);
      } catch (error) {
        // Recording is best-effort: a user with read-only access to the site
        // must still be able to check a file. The result stays on screen and
        // the failure is a note, not an error.
        setNotice(
          `The check completed, but it could not be recorded to this site: ` +
          `${error instanceof Error ? error.message : String(error)}`
        );
      }
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error));
    }
  }, [services, setBusy, setError, setNotice, setRun]);

  const checkLibraryFile = useCallback(async (file: LibraryFile): Promise<void> => {
    setError(undefined);
    setBusy(true, `Downloading ${file.name}`);
    try {
      const source = await services.files.openFile(file);
      await check(source);
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error));
    }
  }, [check, services, setBusy, setError]);

  const checkLocalFile = useCallback(
    (file: File): Promise<void> => check(fromLocalFile(file)),
    [check]
  );

  const exportReport = useCallback(async (): Promise<void> => {
    const run = useCheckerStore.getState().run;
    if (!run) { return; }

    setBusy(true, 'Building the report');
    try {
      const report = await buildReport(run);
      const url = await services.files.uploadFile(
        services.reportFolder, report.fileName, report.buffer
      );
      setBusy(false);
      // The path, not just "saved": on a site with several libraries the user
      // otherwise has to go looking for their own report.
      setNotice(`Report saved to ${url}`);
    } catch (error) {
      setError(
        `The report could not be saved: ` +
        `${error instanceof Error ? error.message : String(error)}`
      );
    }
  }, [services, setBusy, setError, setNotice]);

  return { checkLibraryFile, checkLocalFile, exportReport };
}
