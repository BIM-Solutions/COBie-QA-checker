/**
 * Turning what someone typed in the property pane into a server-relative folder
 * path.
 *
 * Split out of `SpFileService` so it can be tested. The service imports
 * `@microsoft/sp-http`, which does not resolve under jest, and this is pure
 * string handling with a genuine failure mode — so it lives where a test can
 * reach it rather than being verified by hand.
 */

/**
 * Accepts `Shared Documents`, `/sites/X/Shared Documents` or a full absolute
 * URL. Web parts are configured by hand and people type all three; rejecting
 * two of them would be a support burden for no gain.
 *
 * Always returns a decoded path with no trailing slash, because every caller
 * passes the result to SharePoint's `decodedurl=`, which wants the literal
 * folder name.
 */
export function resolveFolderPath(configured: string, siteServerRelativeUrl: string): string {
  const site = siteServerRelativeUrl.replace(/\/+$/, '');
  const value = (configured || '').trim();

  if (value === '') { return `${site}/Shared Documents`; }

  if (value.indexOf('://') !== -1) {
    // Decoded, because `pathname` percent-encodes and the caller's `encode()`
    // deliberately does not. Pasting the browser's own URL for a library yields
    // "Shared%20Documents", and handing that to `decodedurl=` asks SharePoint
    // for a folder whose name literally contains "%20" — a 404 on every
    // library with a space in its name, which is most of them.
    return decodeURIComponent(new URL(value).pathname).replace(/\/+$/, '');
  }

  if (value.charAt(0) === '/') { return value.replace(/\/+$/, ''); }

  return `${site}/${value.replace(/^\/+|\/+$/g, '')}`;
}
