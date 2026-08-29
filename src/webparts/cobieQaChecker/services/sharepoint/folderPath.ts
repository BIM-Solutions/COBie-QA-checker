/**
 * Turning what someone typed in the property pane into a server-relative folder
 * path.
 *
 * Split out of `SpFileService` so it can be tested. The service imports
 * `@microsoft/sp-http`, which does not resolve under jest, and this is pure
 * string handling with a genuine failure mode — so it lives where a test can
 * reach it rather than being verified by hand.
 *
 * Nothing here throws. The input is free text from a property pane, and this
 * runs inside `onInit`: an exception would not surface as a bad path, it would
 * leave the web part blank with nothing on screen to say why. Every branch
 * degrades to a path that is wrong-but-legible instead, so the failure arrives
 * as "could not list files in /sites/X/https:", which names its own cause.
 */

/**
 * Accepts `Shared Documents`, `/sites/X/Shared Documents` or a full absolute
 * URL. Web parts are configured by hand and people type all three; rejecting
 * two of them would be a support burden for no gain.
 *
 * Percent-escapes are decoded **only** for the absolute-URL form, where the
 * encoding is structural — `URL` puts it there. A path typed or pasted directly
 * is taken literally, because there the `%` is a character in a folder name,
 * and decoding it would corrupt every library called something like
 * "100% Complete". Trailing slashes are stripped in all forms.
 *
 * The result is what SharePoint's `decodedurl=` wants: a literal path.
 */
export function resolveFolderPath(configured: string, siteServerRelativeUrl: string): string {
  const site = siteServerRelativeUrl.replace(/\/+$/, '');
  const value = (configured || '').trim();

  if (value === '') { return `${site}/Shared Documents`; }

  if (value.indexOf('://') !== -1) {
    const path = absolutePathOf(value);
    // `undefined` means it looked like a URL and would not parse — a half-typed
    // "https://" or a stray space in the host. Falling through treats it as a
    // path, which produces a visibly wrong folder rather than an exception.
    if (path !== undefined) { return path; }
  }

  if (value.charAt(0) === '/') { return value.replace(/\/+$/, ''); }

  return `${site}/${value.replace(/^\/+|\/+$/g, '')}`;
}

/** The decoded path of an absolute URL, or undefined if it will not parse. */
function absolutePathOf(value: string): string | undefined {
  let pathname: string;
  try {
    pathname = new URL(value).pathname;
  } catch {
    return undefined;
  }
  return decodePath(pathname).replace(/\/+$/, '');
}

/**
 * Decodes percent-escapes, salvaging what it can from input that is not wholly
 * valid.
 *
 * The whole-string decode is tried first because it is the only one that gets
 * multi-byte characters right: `é` is `%C3%A9`, two escapes that mean nothing
 * apart.
 *
 * It throws on a folder named "100% Complete", which reaches here as
 * "100%%20Complete" — a literal `%` the URL constructor left alone, followed by
 * a real escape. The per-escape pass then recovers "100% Complete" exactly,
 * where returning the raw pathname would have left a visible "%20" in the name.
 */
function decodePath(path: string): string {
  try {
    return decodeURIComponent(path);
  } catch {
    return path.replace(/%[0-9A-Fa-f]{2}/g, (escape) => {
      try {
        return decodeURIComponent(escape);
      } catch {
        // Half of a multi-byte pair whose partner is malformed. Left as-is:
        // an escape on screen beats a mangled character.
        return escape;
      }
    });
  }
}
