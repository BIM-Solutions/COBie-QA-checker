import { resolveFolderPath } from './folderPath';

const SITE = '/sites/Project';

describe('resolveFolderPath', () => {
  it('defaults to Shared Documents on this site', () => {
    expect(resolveFolderPath('', SITE)).toBe('/sites/Project/Shared Documents');
  });

  it('treats a bare name as relative to the site', () => {
    expect(resolveFolderPath('COBie', SITE)).toBe('/sites/Project/COBie');
  });

  it('takes a server-relative path as given', () => {
    expect(resolveFolderPath('/sites/Other/Shared Documents', SITE))
      .toBe('/sites/Other/Shared Documents');
  });

  it('decodes an absolute URL rather than leaving it percent-encoded', () => {
    // Pasting the browser's own URL for a library is the obvious thing to do,
    // and `pathname` hands back "Shared%20Documents". The caller's `encode()`
    // deliberately does not percent-encode, so passing that through asks
    // SharePoint for a folder whose name literally contains "%20" - a 404 on
    // every library with a space in its name, which is most of them.
    expect(resolveFolderPath('https://contoso.sharepoint.com/sites/Project/Shared Documents', SITE))
      .toBe('/sites/Project/Shared Documents');
  });

  it('decodes an absolute URL that arrives already encoded', () => {
    expect(resolveFolderPath('https://contoso.sharepoint.com/sites/Project/Shared%20Documents', SITE))
      .toBe('/sites/Project/Shared Documents');
  });

  it('strips a trailing slash from an absolute URL, as the other branches do', () => {
    expect(resolveFolderPath('https://contoso.sharepoint.com/sites/Project/COBie/', SITE))
      .toBe('/sites/Project/COBie');
  });

  it('strips a trailing slash from a server-relative path', () => {
    expect(resolveFolderPath('/sites/Project/COBie/', SITE)).toBe('/sites/Project/COBie');
  });

  it('handles a nested folder inside a library', () => {
    expect(resolveFolderPath('Shared Documents/COBie/Issued', SITE))
      .toBe('/sites/Project/Shared Documents/COBie/Issued');
  });

  it('does not double the separator on a root-site web', () => {
    // `serverRelativeUrl` is "/" on a root site; naive concatenation yields
    // "//Shared Documents".
    expect(resolveFolderPath('COBie', '/')).toBe('/COBie');
    expect(resolveFolderPath('', '/')).toBe('/Shared Documents');
  });

  it('ignores surrounding whitespace', () => {
    expect(resolveFolderPath('  COBie  ', SITE)).toBe('/sites/Project/COBie');
  });
});

describe('resolveFolderPath does not throw on bad input', () => {
  // This runs inside the web part's onInit. An exception here does not surface
  // as a bad path - it leaves the web part blank with nothing to say why.

  it('decodes a multi-byte character in an absolute URL', () => {
    expect(resolveFolderPath('https://contoso.sharepoint.com/sites/Project/Café', SITE))
      .toBe('/sites/Project/Café');
  });

  it('survives a library name containing a literal percent sign', () => {
    // "100% Complete" reaches the decoder as "100%%20Complete": a literal % the
    // URL constructor left alone, then a real escape. A whole-string decode
    // throws URIError on it, which without the salvage pass would blank the
    // web part on a perfectly ordinary library name.
    expect(resolveFolderPath('https://contoso.sharepoint.com/sites/Project/100% Complete', SITE))
      .toBe('/sites/Project/100% Complete');
  });

  it('falls back to a path when a URL-ish value will not parse', () => {
    // Half-typed property-pane values. The result is wrong but legible, and the
    // list call reports it by name.
    expect(() => resolveFolderPath('https://', SITE)).not.toThrow();
    expect(() => resolveFolderPath('://', SITE)).not.toThrow();
    expect(() => resolveFolderPath('ht tp://bad host/x', SITE)).not.toThrow();
  });

  it('does not decode a path typed directly', () => {
    // Only the absolute-URL form is decoded, where the encoding is structural.
    // A typed path is literal, so a folder actually named "50%25" keeps its
    // name instead of silently becoming "50%".
    expect(resolveFolderPath('/sites/Project/50%25', SITE)).toBe('/sites/Project/50%25');
    expect(resolveFolderPath('50%25', SITE)).toBe('/sites/Project/50%25');
  });
});
