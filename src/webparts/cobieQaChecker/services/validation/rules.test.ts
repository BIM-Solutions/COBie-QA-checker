import { buildIndex } from './workbookIndex';
import {
  checkColumns,
  checkCompleteness,
  checkDynamicReferences,
  checkFormats,
  checkPickLists,
  checkReferences,
  checkStructure,
  checkUniqueness
} from './rules';
import type { Finding } from '../../models/findings';
import type { ParsedWorkbook } from '../../models/workbook';
import { CONTACT, cleanWorkbook, sheet, workbook } from './fixtures';

function ruleIds(findings: readonly Finding[]): string[] {
  return findings.map((f) => f.ruleId);
}

function index(wb: ParsedWorkbook): ReturnType<typeof buildIndex> {
  return buildIndex(wb);
}

describe('checkStructure', () => {
  it('reports a missing required sheet as an error', () => {
    const findings = checkStructure(index(workbook([CONTACT])));
    const facility = findings.filter((f) => f.sheet === 'Facility')[0];
    expect(facility.ruleId).toBe('structure.missingSheet');
    expect(facility.severity).toBe('error');
  });

  it('reports a missing expected sheet as a warning, not an error', () => {
    // Zone and System are "expected": their absence is worth saying, but it
    // must not fail an otherwise valid deliverable.
    const findings = checkStructure(index(cleanWorkbook()));
    const zone = findings.filter((f) => f.sheet === 'Zone')[0];
    expect(zone.severity).toBe('warning');
  });

  it('says nothing about an absent optional sheet', () => {
    const findings = checkStructure(index(cleanWorkbook()));
    expect(findings.filter((f) => f.sheet === 'Connection')).toHaveLength(0);
  });

  it('reports a present but empty required sheet', () => {
    const wb = workbook([CONTACT, sheet('Facility', [['Name', 'CreatedBy']])]);
    const findings = checkStructure(index(wb));
    expect(ruleIds(findings.filter((f) => f.sheet === 'Facility'))).toEqual(['structure.emptySheet']);
  });

  it('notes an unrecognised sheet as info only', () => {
    const wb = workbook([CONTACT, sheet('Our Notes', [['Name', 'Value'], ['x', 'y']])]);
    const finding = checkStructure(index(wb)).filter((f) => f.sheet === 'Our Notes')[0];
    expect(finding.severity).toBe('info');
  });
});

describe('checkColumns', () => {
  it('reports a missing required column on a sheet that is present', () => {
    const wb = workbook([sheet('Floor', [['Name', 'CreatedOn'], ['L0', '2026-01-15T09:30:00']])]);
    const findings = checkColumns(index(wb)).filter((f) => f.column === 'CreatedBy');
    expect(findings[0].severity).toBe('error');
  });

  it('accepts an alias spelling rather than calling the column missing', () => {
    // The ExtSystem / ExternalSystem split is real and common. Reporting it
    // would put three spurious findings on every sheet of an otherwise fine file.
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'ExternalSystem'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor', 'Revit']
    ])]);
    expect(checkColumns(index(wb)).filter((f) => f.column === 'ExtSystem')).toHaveLength(0);
  });

  it('matches headers regardless of case and spacing', () => {
    const wb = workbook([sheet('Floor', [
      ['name', 'Created By', 'CREATEDON', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor']
    ])]);
    expect(checkColumns(index(wb)).filter((f) => f.severity === 'error')).toHaveLength(0);
  });
});

describe('checkCompleteness', () => {
  it('reports an empty required cell', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', '']
    ])]);
    const finding = checkCompleteness(index(wb)).filter((f) => f.column === 'Category')[0];
    expect(finding.ruleId).toBe('completeness.missingValue');
    expect(finding.row).toBe(2);
  });

  it('separates a placeholder from a blank', () => {
    // The distinction is load-bearing: a blank is an unconfigured exporter, an
    // "n/a" is a person who declined to fill the field. Different owners,
    // different fixes, so they must not collapse into one finding type.
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'n/a']
    ])]);
    const finding = checkCompleteness(index(wb)).filter((f) => f.column === 'Category')[0];
    expect(finding.ruleId).toBe('placeholder.value');
    expect(finding.value).toBe('n/a');
  });

  it('treats a range of placeholder spellings as empty', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'TBC'],
      ['L1', 'a@b.com', '2026-01-15T09:30:00', '-'],
      ['L2', 'a@b.com', '2026-01-15T09:30:00', 'Unknown']
    ])]);
    const placeholders = checkCompleteness(index(wb))
      .filter((f) => f.ruleId === 'placeholder.value' && f.column === 'Category');
    expect(placeholders).toHaveLength(3);
  });

  it('stays silent on an optional column left empty', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor', 'Ground']
    ])]);
    expect(checkCompleteness(index(wb)).filter((f) => f.sheet === 'Floor')).toHaveLength(0);
  });
});

describe('checkUniqueness', () => {
  it('reports the later duplicate and points at the first', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor']
    ])]);
    const findings = checkUniqueness(index(wb));
    expect(findings).toHaveLength(1);
    expect(findings[0].row).toBe(3);
    expect(findings[0].message).toContain('row 2');
  });

  it('treats keys case-insensitively', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor'],
      ['l0', 'a@b.com', '2026-01-15T09:30:00', 'Floor']
    ])]);
    expect(checkUniqueness(index(wb))).toHaveLength(1);
  });

  it('allows an Attribute name to repeat against different rows', () => {
    // This is the whole point of the Attribute sheet. Keying it on Name alone
    // would report a duplicate for every attribute used more than once, which
    // on a real file is essentially all of them.
    const wb = workbook([sheet('Attribute', [
      ['Name', 'CreatedBy', 'CreatedOn', 'SheetName', 'RowName', 'Value'],
      ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'AHU-01-001', '100'],
      ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'AHU-01-002', '120']
    ])]);
    expect(checkUniqueness(index(wb))).toHaveLength(0);
  });

  it('still reports an Attribute repeated against the same row', () => {
    const wb = workbook([sheet('Attribute', [
      ['Name', 'CreatedBy', 'CreatedOn', 'SheetName', 'RowName', 'Value'],
      ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'AHU-01-001', '100'],
      ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'AHU-01-001', '120']
    ])]);
    expect(checkUniqueness(index(wb))).toHaveLength(1);
  });

  it('does not treat two placeholder names as duplicates of each other', () => {
    // Both rows are already reported as missing a Name. Adding a duplicate
    // finding on top would double-count one defect.
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['n/a', 'a@b.com', '2026-01-15T09:30:00', 'Floor'],
      ['n/a', 'a@b.com', '2026-01-15T09:30:00', 'Floor']
    ])]);
    expect(checkUniqueness(index(wb))).toHaveLength(0);
  });
});

describe('checkReferences', () => {
  it('reports a Space pointing at a Floor that does not exist', () => {
    const wb = cleanWorkbook();
    const broken = workbook(
      wb.sheets.map((s) => s.name === 'Space'
        ? sheet('Space', [
          ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'FloorName', 'Description'],
          ['101', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'Level 99', 'Meeting room']
        ])
        : s)
    );
    const finding = checkReferences(index(broken)).filter((f) => f.column === 'FloorName')[0];
    expect(finding.ruleId).toBe('reference.broken');
    expect(finding.value).toBe('Level 99');
  });

  it('resolves every member of a list reference', () => {
    const wb = workbook([
      CONTACT,
      sheet('Space', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'FloorName', 'Description'],
        ['101', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'L0', 'A'],
        ['102', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'L0', 'B']
      ]),
      sheet('Zone', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'SpaceNames'],
        ['Z1', 'a@b.com', '2026-01-15T09:30:00', 'Fire', '101,102,999']
      ])
    ]);
    const broken = checkReferences(index(wb)).filter((f) => f.column === 'SpaceNames');
    expect(broken).toHaveLength(1);
    expect(broken[0].value).toBe('999');
  });

  it('reports a missing target sheet once, not once per row', () => {
    // A file with no Type sheet would otherwise emit one finding per Component.
    // On a 40,000-row file that is the difference between a usable report and
    // an unusable one.
    const wb = workbook([
      CONTACT,
      sheet('Component', [
        ['Name', 'CreatedBy', 'CreatedOn', 'TypeName', 'Space', 'Description'],
        ['C1', 'a@b.com', '2026-01-15T09:30:00', 'AHU-01', '101', 'x'],
        ['C2', 'a@b.com', '2026-01-15T09:30:00', 'AHU-01', '101', 'x'],
        ['C3', 'a@b.com', '2026-01-15T09:30:00', 'AHU-01', '101', 'x']
      ])
    ]);
    const findings = checkReferences(index(wb)).filter((f) => f.column === 'TypeName');
    expect(findings).toHaveLength(1);
    expect(findings[0].ruleId).toBe('reference.missingTargetSheet');
  });

  it('does not report a broken reference for a blank cell', () => {
    // That is the completeness rule's finding. Reporting both would count one
    // empty cell twice and inflate the error total.
    const wb = workbook([
      CONTACT,
      sheet('Floor', [['Name', 'CreatedBy', 'CreatedOn', 'Category'],
        ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor']]),
      sheet('Space', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'FloorName', 'Description'],
        ['101', 'a@b.com', '2026-01-15T09:30:00', 'Office', '', 'A']
      ])
    ]);
    expect(checkReferences(index(wb)).filter((f) => f.column === 'FloorName')).toHaveLength(0);
  });

  it('does not let a row named "n/a" satisfy a reference to "n/a"', () => {
    const wb = workbook([
      CONTACT,
      sheet('Floor', [['Name', 'CreatedBy', 'CreatedOn', 'Category'],
        ['n/a', 'a@b.com', '2026-01-15T09:30:00', 'Floor']]),
      sheet('Space', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'FloorName', 'Description'],
        ['101', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'Level 00', 'A']
      ])
    ]);
    const findings = checkReferences(index(wb)).filter((f) => f.column === 'FloorName');
    expect(findings).toHaveLength(1);
  });

  it('reports a Manufacturer that is a company name rather than a contact', () => {
    const wb = workbook([
      CONTACT,
      sheet('Type', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description', 'AssetType',
          'Manufacturer', 'ModelNumber', 'WarrantyGuarantorParts', 'WarrantyDurationParts',
          'WarrantyGuarantorLabor', 'WarrantyDurationLabor', 'WarrantyDurationUnit'],
        ['AHU-01', 'a@b.com', '2026-01-15T09:30:00', 'Plant', 'AHU', 'Fixed',
          'Acme Ltd', 'M-1', 'a@b.com', '2', 'a@b.com', '1', 'year']
      ])
    ]);
    const findings = checkReferences(index(wb)).filter((f) => f.column === 'Manufacturer');
    expect(findings[0].value).toBe('Acme Ltd');
  });
});

describe('checkDynamicReferences', () => {
  it('reports an Attribute pointing at a row that does not exist', () => {
    const wb = workbook([
      CONTACT,
      sheet('Component', [
        ['Name', 'CreatedBy', 'CreatedOn', 'TypeName', 'Space', 'Description'],
        ['C1', 'a@b.com', '2026-01-15T09:30:00', 'T1', '101', 'x']
      ]),
      sheet('Attribute', [
        ['Name', 'CreatedBy', 'CreatedOn', 'SheetName', 'RowName', 'Value'],
        ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'C9', '100']
      ])
    ]);
    const findings = checkDynamicReferences(index(wb)).filter((f) => f.ruleId === 'reference.brokenDynamic');
    expect(findings[0].value).toBe('C9');
  });

  it('reports a SheetName naming something that is not a COBie sheet here', () => {
    const wb = workbook([
      CONTACT,
      sheet('Attribute', [
        ['Name', 'CreatedBy', 'CreatedOn', 'SheetName', 'RowName', 'Value'],
        ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Widgets', 'C1', '100']
      ])
    ]);
    const findings = checkDynamicReferences(index(wb));
    expect(findings[0].ruleId).toBe('reference.unknownSheetName');
    expect(findings[0].value).toBe('Widgets');
  });

  it('resolves a valid dynamic reference quietly', () => {
    const wb = workbook([
      CONTACT,
      sheet('Component', [
        ['Name', 'CreatedBy', 'CreatedOn', 'TypeName', 'Space', 'Description'],
        ['C1', 'a@b.com', '2026-01-15T09:30:00', 'T1', '101', 'x']
      ]),
      sheet('Attribute', [
        ['Name', 'CreatedBy', 'CreatedOn', 'SheetName', 'RowName', 'Value'],
        ['Airflow', 'a@b.com', '2026-01-15T09:30:00', 'Component', 'C1', '100']
      ])
    ]);
    expect(checkDynamicReferences(index(wb))).toHaveLength(0);
  });
});

describe('checkFormats', () => {
  it('reports a non-ISO date', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '15/01/2026', 'Floor']
    ])]);
    const finding = checkFormats(index(wb)).filter((f) => f.column === 'CreatedOn')[0];
    expect(finding.ruleId).toBe('format.isoDateTime');
    expect(finding.severity).toBe('error');
  });

  it('rejects a well-formed date that is not a real day', () => {
    // 2026-02-30 passes the pattern. Constructing it through `Date` would
    // silently roll it into March and report the file as clean.
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-02-30T09:00:00', 'Floor']
    ])]);
    expect(checkFormats(index(wb)).filter((f) => f.column === 'CreatedOn')).toHaveLength(1);
  });

  it('accepts 29 February in a leap year', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2024-02-29T09:00:00', 'Floor']
    ])]);
    expect(checkFormats(index(wb)).filter((f) => f.column === 'CreatedOn')).toHaveLength(0);
  });

  it('reports a non-email in an email column', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'Ann Bell', '2026-01-15T09:30:00', 'Floor']
    ])]);
    expect(checkFormats(index(wb)).filter((f) => f.ruleId === 'format.email')).toHaveLength(1);
  });

  it('reports text in a number column', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Elevation'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor', 'ground level']
    ])]);
    const finding = checkFormats(index(wb)).filter((f) => f.column === 'Elevation')[0];
    expect(finding.ruleId).toBe('format.number');
    // Elevation is "expected", not "required", so a bad value must not fail the file.
    expect(finding.severity).toBe('warning');
  });

  it('accepts negative and decimal numbers', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Elevation'],
      ['B1', 'a@b.com', '2026-01-15T09:30:00', 'Floor', '-3500.5']
    ])]);
    expect(checkFormats(index(wb)).filter((f) => f.column === 'Elevation')).toHaveLength(0);
  });

  it('does not report a format defect for a placeholder', () => {
    // Already reported as a placeholder. Reporting it again as a bad number
    // would count one cell twice.
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Elevation'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Floor', 'n/a']
    ])]);
    expect(checkFormats(index(wb)).filter((f) => f.column === 'Elevation')).toHaveLength(0);
  });
});

describe('checkPickLists', () => {
  it('reports a value outside a COBie-fixed list', () => {
    const wb = workbook([
      CONTACT,
      sheet('Type', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description', 'AssetType',
          'Manufacturer', 'ModelNumber', 'WarrantyGuarantorParts', 'WarrantyDurationParts',
          'WarrantyGuarantorLabor', 'WarrantyDurationLabor', 'WarrantyDurationUnit'],
        ['T1', 'a@b.com', '2026-01-15T09:30:00', 'Plant', 'x', 'Portable',
          'a@b.com', 'M1', 'a@b.com', '1', 'a@b.com', '1', 'year']
      ])
    ]);
    const finding = checkPickLists(index(wb)).filter((f) => f.column === 'AssetType')[0];
    expect(finding.value).toBe('Portable');
    // Never an error: a project may legitimately extend an enumeration.
    expect(finding.severity).toBe('warning');
  });

  it('accepts a fixed-list value regardless of case', () => {
    const wb = workbook([
      CONTACT,
      sheet('Type', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description', 'AssetType',
          'Manufacturer', 'ModelNumber', 'WarrantyGuarantorParts', 'WarrantyDurationParts',
          'WarrantyGuarantorLabor', 'WarrantyDurationLabor', 'WarrantyDurationUnit'],
        ['T1', 'a@b.com', '2026-01-15T09:30:00', 'Plant', 'x', 'FIXED',
          'a@b.com', 'M1', 'a@b.com', '1', 'a@b.com', '1', 'year']
      ])
    ]);
    expect(checkPickLists(index(wb)).filter((f) => f.column === 'AssetType')).toHaveLength(0);
  });

  it('prefers the file own PickLists sheet over the built-in list', () => {
    // A project that ships its own enumerations is entitled to extend them.
    // Checking against the built-ins anyway would flag every correctly
    // classified row.
    const wb = workbook([
      CONTACT,
      sheet('PickLists', [['AssetType'], ['Fixed'], ['Moveable'], ['Portable']]),
      sheet('Type', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description', 'AssetType',
          'Manufacturer', 'ModelNumber', 'WarrantyGuarantorParts', 'WarrantyDurationParts',
          'WarrantyGuarantorLabor', 'WarrantyDurationLabor', 'WarrantyDurationUnit'],
        ['T1', 'a@b.com', '2026-01-15T09:30:00', 'Plant', 'x', 'Portable',
          'a@b.com', 'M1', 'a@b.com', '1', 'a@b.com', '1', 'year']
      ])
    ]);
    expect(checkPickLists(index(wb)).filter((f) => f.column === 'AssetType')).toHaveLength(0);
  });

  it('says nothing about a Category with no list to check against', () => {
    // Classification is a project decision. Without a PickLists sheet there is
    // no authority to check Category against, and inventing one would report a
    // defect on every properly classified file.
    const wb = workbook([CONTACT, sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'Ss_25_10_30 Floors']
    ])]);
    expect(checkPickLists(index(wb)).filter((f) => f.column === 'Category')).toHaveLength(0);
  });
});
