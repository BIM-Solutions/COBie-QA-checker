import { formatLocalIso, sheetFromGrid } from './workbookReader';

describe('sheetFromGrid', () => {
  it('anchors rows to their real spreadsheet row number', () => {
    // The header is row 1, so the first record is row 2. Every finding quotes
    // this number and a user goes to that row in Excel; an off-by-one makes the
    // whole report untrustworthy.
    const sheet = sheetFromGrid('Floor', [
      ['Name', 'CreatedBy'],
      ['Level 00', 'a@b.com'],
      ['Level 01', 'a@b.com']
    ]);
    expect(sheet.rows.map((r) => r.sourceRow)).toEqual([2, 3]);
  });

  it('keys cells on the normalised header', () => {
    const sheet = sheetFromGrid('Floor', [['Ext System', 'Name'], ['Revit', 'L0']]);
    expect(sheet.rows[0].cells.extsystem).toBe('Revit');
    expect(sheet.rows[0].cells.name).toBe('L0');
  });

  it('skips wholly blank rows without consuming a row number', () => {
    // COBie templates ship with hundreds of blank padding rows. Keeping them
    // would bury every real finding under one "missing Name" per empty row —
    // and the surviving rows must still report their true position.
    const sheet = sheetFromGrid('Floor', [
      ['Name'],
      ['L0'],
      ['   '],
      [''],
      ['L1']
    ]);
    expect(sheet.rows.map((r) => r.sourceRow)).toEqual([2, 5]);
  });

  it('finds a header row sitting below a title band', () => {
    const sheet = sheetFromGrid('Space', [
      ['COBie export — Project X', ''],
      [''],
      ['Name', 'FloorName'],
      ['101', 'L0']
    ]);
    expect(sheet.headers).toEqual(['Name', 'FloorName']);
    expect(sheet.rows[0].sourceRow).toBe(4);
  });

  it('reads a single-column sheet rather than dropping every row', () => {
    // Requiring two filled cells to call a row a header rejected this outright,
    // and a sheet read as headerless loses all its rows silently - the worst
    // way for a checker to be wrong.
    const sheet = sheetFromGrid('Floor', [['Name'], ['L0'], ['L1']]);
    expect(sheet.headers).toEqual(['Name']);
    expect(sheet.rows.map((r) => r.cells.name)).toEqual(['L0', 'L1']);
  });

  it('reports no rows when there is no header at all', () => {
    expect(sheetFromGrid('Empty', []).rows).toHaveLength(0);
    expect(sheetFromGrid('Empty', [[]]).headers).toHaveLength(0);
  });

  it('ignores cells under an empty header', () => {
    const sheet = sheetFromGrid('Floor', [['Name', '', 'Category'], ['L0', 'junk', 'Floor']]);
    expect(sheet.rows[0].cells).toEqual({ name: 'L0', category: 'Floor' });
  });
});

describe('formatLocalIso', () => {
  it('uses the date own local fields rather than shifting to UTC', () => {
    // Built from local components, so a machine west of UTC would report the
    // previous day if this went via toISOString.
    const date = new Date(2026, 0, 15, 9, 30, 0);
    expect(formatLocalIso(date)).toBe('2026-01-15T09:30:00');
  });

  it('pads every component', () => {
    expect(formatLocalIso(new Date(2026, 8, 5, 4, 3, 2))).toBe('2026-09-05T04:03:02');
  });
});
