import { SPHttpClient } from '@microsoft/sp-http';
import type { CheckRun } from '../../models/findings';

/**
 * Run history, kept in a SharePoint list on the site the web part sits on.
 *
 * Everything stays inside Microsoft 365: no Azure, no database, no third-party
 * service. That constraint is the reason the checker is worth building as a web
 * part at all, and it is why history is a list rather than anything cleverer.
 *
 * What is stored is the *summary*, not the findings. A single run on a real
 * deliverable produces tens of thousands of findings; writing those to a list
 * would blow past the 5,000-item view threshold on the first check and make the
 * list unusable for the thing it is actually good at, which is showing whether
 * this week's export is better than last week's. Findings live in the exported
 * report.
 */

export interface RunHistoryEntry {
  readonly id: number;
  readonly fileName: string;
  readonly checkedOn: string;
  readonly checkedBy: string;
  readonly errors: number;
  readonly warnings: number;
  readonly completeness: number;
  readonly passed: boolean;
}

export const RUN_HISTORY_LIST = 'COBie Check History';

/** Internal names are fixed at creation and never change; titles can. */
const FIELDS = [
  { name: 'CobieFileName', type: 'Text', title: 'File' },
  { name: 'CobieCheckedOn', type: 'DateTime', title: 'Checked on' },
  { name: 'CobieCheckedBy', type: 'Text', title: 'Checked by' },
  { name: 'CobieErrors', type: 'Number', title: 'Errors' },
  { name: 'CobieWarnings', type: 'Number', title: 'Warnings' },
  { name: 'CobieCompleteness', type: 'Number', title: 'Completeness' },
  { name: 'CobiePassed', type: 'Boolean', title: 'Passed' }
];

const FIELD_TYPE_IDS: Readonly<Record<string, number>> = {
  Text: 2,
  Number: 9,
  DateTime: 4,
  Boolean: 8
};

interface ListItem {
  Id: number;
  CobieFileName?: string;
  CobieCheckedOn?: string;
  CobieCheckedBy?: string;
  CobieErrors?: number;
  CobieWarnings?: number;
  CobieCompleteness?: number;
  CobiePassed?: boolean;
}

export class SpRunHistoryService {
  /**
   * Cached so the checker does not re-probe the list on every run. Undefined
   * means "not yet asked", which is different from false.
   */
  private _exists: boolean | undefined;

  public constructor(
    private readonly _client: SPHttpClient,
    private readonly _webAbsoluteUrl: string
  ) {}

  public async listExists(): Promise<boolean> {
    if (this._exists !== undefined) { return this._exists; }

    const response = await this._client.get(
      `${this._webAbsoluteUrl}/_api/web/lists/getByTitle('${RUN_HISTORY_LIST}')?$select=Id`,
      SPHttpClient.configurations.v1
    );
    this._exists = response.ok;
    return this._exists;
  }

  /**
   * Creates the list and its columns.
   *
   * The site provisions itself rather than needing PnP.PowerShell and a tenant
   * admin, so someone can drop the web part on a site and use it. Adding a
   * field that already exists is treated as success: a half-completed earlier
   * attempt must be recoverable by running this again.
   */
  public async ensureList(): Promise<void> {
    if (await this.listExists()) { return; }

    const create = await this._client.post(
      `${this._webAbsoluteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata', Accept: 'application/json;odata=nometadata' },
        body: JSON.stringify({
          BaseTemplate: 100,
          Title: RUN_HISTORY_LIST,
          Description: 'One row per COBie check run by the COBie QA Checker web part.'
        })
      }
    );

    // 409 means someone else created it between the probe and the call - two
    // people opening the web part at once on a fresh site. That is success, not
    // a failure.
    if (!create.ok && create.status !== 409) {
      throw new Error(`Could not create the "${RUN_HISTORY_LIST}" list: ${create.status} ${create.statusText}`);
    }

    for (let i = 0; i < FIELDS.length; i++) {
      await this._addField(FIELDS[i]);
    }

    this._exists = true;
  }

  private async _addField(field: { name: string; type: string; title: string }): Promise<void> {
    const response = await this._client.post(
      `${this._webAbsoluteUrl}/_api/web/lists/getByTitle('${RUN_HISTORY_LIST}')/fields`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata', Accept: 'application/json;odata=nometadata' },
        body: JSON.stringify({
          Title: field.name,
          FieldTypeKind: FIELD_TYPE_IDS[field.type],
          StaticName: field.name
        })
      }
    );

    if (!response.ok && response.status !== 409) {
      throw new Error(`Could not add the ${field.name} column: ${response.status} ${response.statusText}`);
    }
  }

  /**
   * Records a run. Best-effort by design: a site where the user cannot write to
   * the list must still be able to *check* a file, so the caller treats a
   * rejection here as a note rather than a failure of the check.
   */
  public async recordRun(run: CheckRun): Promise<void> {
    await this.ensureList();

    const response = await this._client.post(
      `${this._webAbsoluteUrl}/_api/web/lists/getByTitle('${RUN_HISTORY_LIST}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Content-Type': 'application/json;odata=nometadata', Accept: 'application/json;odata=nometadata' },
        body: JSON.stringify({
          Title: run.fileName,
          CobieFileName: run.fileName,
          CobieCheckedOn: run.checkedOn,
          CobieCheckedBy: run.checkedBy,
          CobieErrors: run.errorCount,
          CobieWarnings: run.warningCount,
          // Stored as a percentage rather than a fraction: a SharePoint list
          // view shows "82" far more usefully than "0.82".
          CobieCompleteness: Math.round(run.completeness * 100),
          CobiePassed: run.passed
        })
      }
    );

    if (!response.ok) {
      throw new Error(`Could not record the run: ${response.status} ${response.statusText}`);
    }
  }

  /** Most recent runs first. */
  public async recentRuns(top: number): Promise<RunHistoryEntry[]> {
    if (!(await this.listExists())) { return []; }

    const url =
      `${this._webAbsoluteUrl}/_api/web/lists/getByTitle('${RUN_HISTORY_LIST}')/items` +
      `?$select=Id,CobieFileName,CobieCheckedOn,CobieCheckedBy,CobieErrors,CobieWarnings,` +
      `CobieCompleteness,CobiePassed&$orderby=CobieCheckedOn desc&$top=${top}`;

    const response = await this._client.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) { return []; }

    const body = await response.json();
    const items: ListItem[] = (body && body.value) || [];

    return items.map((item) => ({
      id: item.Id,
      fileName: item.CobieFileName || '',
      checkedOn: item.CobieCheckedOn || '',
      checkedBy: item.CobieCheckedBy || '',
      errors: item.CobieErrors || 0,
      warnings: item.CobieWarnings || 0,
      // Back to a fraction, so every consumer in the app deals in one unit.
      completeness: (item.CobieCompleteness || 0) / 100,
      passed: item.CobiePassed === true
    }));
  }
}
