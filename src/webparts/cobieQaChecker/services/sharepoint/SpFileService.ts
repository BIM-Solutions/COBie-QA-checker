import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { FileSource } from '../parse';

/**
 * Reading COBie files out of a SharePoint document library, and writing reports
 * back into one.
 *
 * This is the whole reason the checker is a web part rather than a page: the
 * deliverable already lives in SharePoint, and asking someone to download it,
 * check it elsewhere and upload a report is the workflow this replaces.
 *
 * Uses the REST API through `SPHttpClient` rather than PnPjs. The calls needed
 * here are three, the SPFx client already carries the request digest and auth,
 * and a library the size of PnPjs in the entry bundle would cost more than it
 * saves.
 */

export interface LibraryFile {
  readonly name: string;
  /** Server-relative URL, which is what every other call here takes. */
  readonly serverRelativeUrl: string;
  readonly size: number;
  readonly modified: string;
  readonly modifiedBy?: string;
}

/** Only these reach the checker; anything else in the library is not a COBie file. */
const COBIE_EXTENSIONS = /\.(xlsx|xlsm|xls|csv|tsv|txt)$/i;

interface SpFileItem {
  Name: string;
  ServerRelativeUrl: string;
  Length: string;
  TimeLastModified: string;
  ModifiedBy?: { Title?: string };
}

export class SpFileService {
  public constructor(
    private readonly _client: SPHttpClient,
    private readonly _webAbsoluteUrl: string,
    private readonly _webServerRelativeUrl: string
  ) {}

  /**
   * Files in a library folder, newest first.
   *
   * `$expand=ModifiedBy` costs nothing here and turns "which of these four
   * near-identical exports am I looking at" into a question the list can answer.
   */
  public async listFiles(folderUrl: string): Promise<LibraryFile[]> {
    const folder = this.resolveFolder(folderUrl);
    const url =
      `${this._webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encode(folder)}')` +
      `/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,ModifiedBy/Title` +
      `&$expand=ModifiedBy&$top=500`;

    const response = await this._client.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(await describeFailure(response, `Could not list files in ${folder}`));
    }

    const body = await response.json();
    const items: SpFileItem[] = (body && body.value) || [];

    return items
      .filter((item) => COBIE_EXTENSIONS.test(item.Name))
      .map((item) => ({
        name: item.Name,
        serverRelativeUrl: item.ServerRelativeUrl,
        size: parseInt(item.Length, 10) || 0,
        modified: item.TimeLastModified,
        modifiedBy: item.ModifiedBy && item.ModifiedBy.Title
      }))
      .sort((a, b) => (a.modified < b.modified ? 1 : -1));
  }

  /**
   * Downloads a file as something `readWorkbook` can consume.
   *
   * The buffer is fetched once and shared by both accessors rather than being
   * re-requested per call: a 40MB COBie workbook downloaded twice because the
   * reader asked for text and then bytes is a bad afternoon on a site link.
   */
  public async openFile(file: LibraryFile): Promise<FileSource> {
    const url =
      `${this._webAbsoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='` +
      `${encode(file.serverRelativeUrl)}')/$value`;

    const response = await this._client.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(await describeFailure(response, `Could not download ${file.name}`));
    }

    const buffer = await response.arrayBuffer();

    return {
      name: file.name,
      arrayBuffer: () => Promise.resolve(buffer),
      text: () => Promise.resolve(decodeUtf8(buffer))
    };
  }

  /**
   * Writes a report into a library, overwriting any report of the same name.
   *
   * Overwrite is deliberate. Reports are named after the file and the check
   * time, so a collision means the same check was exported twice, and keeping
   * both would litter the library with duplicates.
   */
  public async uploadFile(folderUrl: string, name: string, content: ArrayBuffer): Promise<string> {
    const folder = this.resolveFolder(folderUrl);
    const url =
      `${this._webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encode(folder)}')` +
      `/Files/AddUsingPath(DecodedUrl='${encode(name)}',Overwrite=true)`;

    const response = await this._client.post(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata' },
      body: content
    });

    if (!response.ok) {
      throw new Error(await describeFailure(response, `Could not save ${name} to ${folder}`));
    }

    const body = await response.json();
    return (body && body.ServerRelativeUrl) || `${folder}/${name}`;
  }

  /**
   * Accepts `Shared Documents`, `/sites/X/Shared Documents` or a full absolute
   * URL. Web parts are configured by hand in the property pane and people type
   * all three; rejecting two of them would be a support burden for no gain.
   */
  public resolveFolder(configured: string): string {
    const value = (configured || '').trim();
    if (value === '') { return `${this.siteRoot()}/Shared Documents`; }
    if (value.indexOf('://') !== -1) { return new URL(value).pathname; }
    if (value.charAt(0) === '/') { return value.replace(/\/+$/, ''); }
    return `${this.siteRoot()}/${value.replace(/^\/+|\/+$/g, '')}`;
  }

  private siteRoot(): string {
    return this._webServerRelativeUrl.replace(/\/+$/, '');
  }
}

/**
 * SharePoint's `decodedurl=` takes a literal path, so the only characters that
 * need escaping are the ones that would terminate the OData string literal.
 * Running it through `encodeURIComponent` instead would double-encode every
 * space in "Shared Documents" and 404.
 */
function encode(path: string): string {
  return path.replace(/'/g, "''");
}

function decodeUtf8(buffer: ArrayBuffer): string {
  // TextDecoder is present in every browser SPFx 1.23 supports; the manual
  // fallback is for the jest environment, where it is not always defined.
  const decoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : undefined;
  if (decoder) { return decoder.decode(buffer); }

  const bytes = new Uint8Array(buffer);
  let text = '';
  for (let i = 0; i < bytes.length; i++) { text += String.fromCharCode(bytes[i]); }
  return decodeURIComponent(escape(text));
}

/**
 * SharePoint puts the useful part of a failure in the body, not the status
 * text. "403" alone sends people to the wrong problem; "the folder does not
 * exist" sends them to the right one.
 */
async function describeFailure(response: SPHttpClientResponse, context: string): Promise<string> {
  let detail = `${response.status} ${response.statusText}`;
  try {
    const body = await response.json();
    const message = body && body.error && body.error.message;
    const value = message && (typeof message === 'string' ? message : message.value);
    if (value) { detail = value; }
  } catch {
    // A non-JSON error body is common enough (an HTML sign-in page, say) that
    // failing to parse it must not replace the real failure with a parse error.
  }
  return `${context}: ${detail}`;
}
