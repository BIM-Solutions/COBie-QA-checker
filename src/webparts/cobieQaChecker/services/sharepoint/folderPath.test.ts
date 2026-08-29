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
