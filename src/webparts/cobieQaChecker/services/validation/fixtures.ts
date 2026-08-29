import type { ParsedSheet, ParsedWorkbook } from '../../models/workbook';
import { sheetFromGrid } from '../parse/workbookReader';

/**
 * Builders for the test workbooks.
 *
 * Written as grids rather than as `ParsedSheet` literals so the tests exercise
 * the same header-matching and row-numbering the real reader produces. A
 * fixture that bypassed `sheetFromGrid` could pass while the parser and the
 * rules disagreed about what a row is.
 */

export function sheet(name: string, grid: string[][]): ParsedSheet {
  return sheetFromGrid(name, grid);
}

export function workbook(sheets: ParsedSheet[], fileName?: string): ParsedWorkbook {
  return { fileName: fileName || 'test.xlsx', sheets, readWarnings: [] };
}

/** A minimal Contact sheet, so `CreatedBy` resolves and does not drown the output. */
export const CONTACT = sheet('Contact', [
  ['Email', 'CreatedBy', 'CreatedOn', 'Company', 'Phone', 'GivenName', 'FamilyName', 'Category',
    'ExtSystem', 'ExtObject', 'ExtIdentifier'],
  ['a@b.com', 'a@b.com', '2026-01-15T09:30:00', 'Acme', '0123', 'Ann', 'Bell', 'Contact',
    'Revit', 'Contact', 'c1']
]);

/**
 * A COBie file that passes every rule.
 *
 * Its job in the suite is to prove the rules stay quiet on good input. Without
 * it, a rule that fires on everything would still satisfy every "reports X"
 * test in the file.
 */
export function cleanWorkbook(): ParsedWorkbook {
  return workbook([
    CONTACT,
    sheet('Facility', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'ProjectName', 'SiteName', 'LinearUnits',
        'AreaUnits', 'VolumeUnits', 'CurrencyUnit'],
      ['HQ', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'P1', 'S1', 'millimeters',
        'square meters', 'cubic meters', 'GBP']
    ]),
    sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['Level 00', 'a@b.com', '2026-01-15T09:30:00', 'Floor']
    ]),
    sheet('Space', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'FloorName', 'Description'],
      ['101', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'Level 00', 'Meeting room']
    ]),
    sheet('Type', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'Description', 'AssetType', 'Manufacturer',
        'ModelNumber', 'WarrantyGuarantorParts', 'WarrantyDurationParts',
        'WarrantyGuarantorLabor', 'WarrantyDurationLabor', 'WarrantyDurationUnit'],
      ['AHU-01', 'a@b.com', '2026-01-15T09:30:00', 'Plant', 'Air handling unit', 'Fixed',
        'a@b.com', 'M-1', 'a@b.com', '2', 'a@b.com', '1', 'year']
    ]),
    sheet('Component', [
      ['Name', 'CreatedBy', 'CreatedOn', 'TypeName', 'Space', 'Description'],
      ['AHU-01-001', 'a@b.com', '2026-01-15T09:30:00', 'AHU-01', '101', 'Unit 1']
    ])
  ]);
}
