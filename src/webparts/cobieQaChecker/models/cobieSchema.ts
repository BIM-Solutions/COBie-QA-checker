/**
 * The COBie 2.4 schema, as the checker understands it.
 *
 * This file is the single source of truth for what a COBie deliverable *should*
 * contain. Every rule in `services/validation` reads it rather than hard-coding
 * sheet or column names, so extending the checker to a different COBie profile
 * (a client's own MIDP, say) is a change here and nowhere else.
 *
 * Column names are the COBie spreadsheet header spellings, not IFC attribute
 * names. Where the published schema and the wild disagree — `ExternalSystem`
 * versus `ExtSystem` is the notorious one — the alternates are listed in
 * `aliases` so a file exported by a tool using the other spelling still
 * validates rather than reporting every column as missing.
 */

/** Sheet names in the order COBie defines them. */
export const COBIE_SHEETS = [
  'Contact',
  'Facility',
  'Floor',
  'Space',
  'Zone',
  'Type',
  'Component',
  'System',
  'Assembly',
  'Connection',
  'Spare',
  'Resource',
  'Job',
  'Impact',
  'Document',
  'Attribute',
  'Coordinate',
  'Issue'
] as const;

export type CobieSheetName = (typeof COBIE_SHEETS)[number];

/**
 * How strictly a field is required.
 *
 * COBie's own wording is "required", "referenced" and "as specified". The
 * checker collapses that to three levels because only the consequence matters:
 * a missing `required` value makes the deliverable non-compliant, a missing
 * `expected` one makes it less useful, and `optional` is never reported.
 */
export type Requirement = 'required' | 'expected' | 'optional';

/** The kind of value a column holds, which decides which format rules apply. */
export type ValueType =
  | 'text'
  | 'email'
  | 'isoDateTime'
  | 'number'
  | 'integer'
  | 'list';

export interface CobieColumn {
  readonly name: string;
  readonly requirement: Requirement;
  readonly type: ValueType;
  /** Alternate header spellings accepted from real-world exporters. */
  readonly aliases?: readonly string[];
  /** Name of the pick list in the PickLists sheet this column draws from. */
  readonly pickList?: string;
}

/**
 * A foreign key from one sheet to another.
 *
 * `list: true` means the cell holds a comma-separated set of keys and every one
 * of them must resolve — that is how COBie models Zone→Space and
 * System→Component membership.
 */
export interface CobieReference {
  readonly column: string;
  readonly targetSheet: CobieSheetName;
  /** Column on the target sheet the value must match. Defaults to the target's key. */
  readonly targetColumn?: string;
  readonly list?: boolean;
  readonly requirement: Requirement;
}

/**
 * A polymorphic reference: the row names a sheet in one column and a row of
 * that sheet in another. Attribute, Document, Coordinate, Impact, Assembly and
 * Connection all work this way, and they are where most real COBie files break,
 * because nothing in a spreadsheet stops the pair from disagreeing.
 */
export interface CobieDynamicReference {
  readonly sheetColumn: string;
  readonly rowColumn: string;
  readonly list?: boolean;
  readonly requirement: Requirement;
}

export interface CobieSheet {
  readonly name: CobieSheetName;
  /** Whether the sheet itself must be present and populated. */
  readonly requirement: Requirement;
  /** Column whose value identifies a row. `Email` on Contact, `Name` elsewhere. */
  readonly key: string;
  readonly columns: readonly CobieColumn[];
  readonly references?: readonly CobieReference[];
  readonly dynamicReferences?: readonly CobieDynamicReference[];
}

/**
 * Columns COBie repeats on every sheet. Spelled out once and spread into each
 * definition rather than merged at read time, so a sheet can still override the
 * requirement level — Facility's `Category`, for instance, is required where
 * Resource's is not.
 */
const CREATED: readonly CobieColumn[] = [
  { name: 'CreatedBy', requirement: 'required', type: 'email' },
  { name: 'CreatedOn', requirement: 'required', type: 'isoDateTime' }
];

const EXTERNAL: readonly CobieColumn[] = [
  { name: 'ExtSystem', requirement: 'expected', type: 'text', aliases: ['ExternalSystem'] },
  { name: 'ExtObject', requirement: 'expected', type: 'text', aliases: ['ExternalObject'] },
  { name: 'ExtIdentifier', requirement: 'expected', type: 'text', aliases: ['ExternalIdentifier'] }
];

/** Every sheet but Facility carries `CreatedBy`, so the rule is declared once. */
const CREATED_BY_IS_A_CONTACT: CobieReference = {
  column: 'CreatedBy',
  targetSheet: 'Contact',
  requirement: 'required'
};

export const COBIE_SCHEMA: Readonly<Record<CobieSheetName, CobieSheet>> = {
  Contact: {
    name: 'Contact',
    requirement: 'required',
    // Contact is the one sheet keyed on something other than Name: every
    // `CreatedBy` in the file is an email address, so email is the identity.
    key: 'Email',
    columns: [
      { name: 'Email', requirement: 'required', type: 'email' },
      ...CREATED,
      { name: 'Category', requirement: 'expected', type: 'text', pickList: 'ContactCategory' },
      { name: 'Company', requirement: 'required', type: 'text' },
      { name: 'Phone', requirement: 'required', type: 'text' },
      ...EXTERNAL,
      { name: 'Department', requirement: 'optional', type: 'text' },
      { name: 'OrganizationCode', requirement: 'optional', type: 'text' },
      { name: 'GivenName', requirement: 'expected', type: 'text' },
      { name: 'FamilyName', requirement: 'expected', type: 'text' },
      { name: 'Street', requirement: 'optional', type: 'text' },
      { name: 'PostalBox', requirement: 'optional', type: 'text' },
      { name: 'Town', requirement: 'optional', type: 'text' },
      { name: 'StateRegion', requirement: 'optional', type: 'text' },
      { name: 'PostalCode', requirement: 'optional', type: 'text' },
      { name: 'Country', requirement: 'optional', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT]
  },

  Facility: {
    name: 'Facility',
    requirement: 'required',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'FacilityCategory' },
      { name: 'ProjectName', requirement: 'required', type: 'text' },
      { name: 'SiteName', requirement: 'required', type: 'text' },
      { name: 'LinearUnits', requirement: 'required', type: 'text', pickList: 'LinearUnit' },
      { name: 'AreaUnits', requirement: 'required', type: 'text', pickList: 'AreaUnit' },
      { name: 'VolumeUnits', requirement: 'required', type: 'text', pickList: 'VolumeUnit' },
      { name: 'CurrencyUnit', requirement: 'required', type: 'text' },
      { name: 'AreaMeasurement', requirement: 'expected', type: 'text' },
      { name: 'ExternalSystem', requirement: 'expected', type: 'text', aliases: ['ExtSystem'] },
      { name: 'ExternalProjectObject', requirement: 'optional', type: 'text' },
      { name: 'ExternalProjectIdentifier', requirement: 'optional', type: 'text' },
      { name: 'ExternalSiteObject', requirement: 'optional', type: 'text' },
      { name: 'ExternalSiteIdentifier', requirement: 'optional', type: 'text' },
      { name: 'ExternalFacilityObject', requirement: 'optional', type: 'text' },
      { name: 'ExternalFacilityIdentifier', requirement: 'optional', type: 'text' },
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'ProjectDescription', requirement: 'optional', type: 'text' },
      { name: 'SiteDescription', requirement: 'optional', type: 'text' },
      { name: 'Phase', requirement: 'expected', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT]
  },

  Floor: {
    name: 'Floor',
    requirement: 'required',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'FloorType' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'Elevation', requirement: 'expected', type: 'number' },
      { name: 'Height', requirement: 'expected', type: 'number' }
    ],
    references: [CREATED_BY_IS_A_CONTACT]
  },

  Space: {
    name: 'Space',
    requirement: 'required',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'SpaceCategory' },
      { name: 'FloorName', requirement: 'required', type: 'text' },
      { name: 'Description', requirement: 'required', type: 'text' },
      ...EXTERNAL,
      { name: 'RoomTag', requirement: 'expected', type: 'text' },
      { name: 'UsableHeight', requirement: 'expected', type: 'number' },
      { name: 'GrossArea', requirement: 'expected', type: 'number' },
      { name: 'NetArea', requirement: 'expected', type: 'number' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'FloorName', targetSheet: 'Floor', requirement: 'required' }
    ]
  },

  Zone: {
    name: 'Zone',
    requirement: 'expected',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'ZoneCategory' },
      { name: 'SpaceNames', requirement: 'required', type: 'list' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'SpaceNames', targetSheet: 'Space', list: true, requirement: 'required' }
    ]
  },

  Type: {
    name: 'Type',
    requirement: 'required',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'TypeCategory' },
      { name: 'Description', requirement: 'required', type: 'text' },
      { name: 'AssetType', requirement: 'required', type: 'text', pickList: 'AssetType' },
      { name: 'Manufacturer', requirement: 'required', type: 'email' },
      { name: 'ModelNumber', requirement: 'required', type: 'text' },
      { name: 'WarrantyGuarantorParts', requirement: 'required', type: 'email' },
      { name: 'WarrantyDurationParts', requirement: 'required', type: 'number' },
      { name: 'WarrantyGuarantorLabor', requirement: 'required', type: 'email' },
      { name: 'WarrantyDurationLabor', requirement: 'required', type: 'number' },
      { name: 'WarrantyDurationUnit', requirement: 'required', type: 'text', pickList: 'DurationUnit' },
      ...EXTERNAL,
      { name: 'ReplacementCost', requirement: 'expected', type: 'number' },
      { name: 'ExpectedLife', requirement: 'expected', type: 'number' },
      { name: 'DurationUnit', requirement: 'expected', type: 'text', pickList: 'DurationUnit' },
      { name: 'WarrantyDescription', requirement: 'optional', type: 'text' },
      { name: 'NominalLength', requirement: 'optional', type: 'number' },
      { name: 'NominalWidth', requirement: 'optional', type: 'number' },
      { name: 'NominalHeight', requirement: 'optional', type: 'number' },
      { name: 'ModelReference', requirement: 'optional', type: 'text' },
      { name: 'Shape', requirement: 'optional', type: 'text' },
      { name: 'Size', requirement: 'optional', type: 'text' },
      { name: 'Color', requirement: 'optional', type: 'text' },
      { name: 'Finish', requirement: 'optional', type: 'text' },
      { name: 'Grade', requirement: 'optional', type: 'text' },
      { name: 'Material', requirement: 'optional', type: 'text' },
      { name: 'Constituents', requirement: 'optional', type: 'text' },
      { name: 'Features', requirement: 'optional', type: 'text' },
      { name: 'AccessibilityPerformance', requirement: 'optional', type: 'text' },
      { name: 'CodePerformance', requirement: 'optional', type: 'text' },
      { name: 'SustainabilityPerformance', requirement: 'optional', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      // Manufacturer and the two warranty guarantors are contacts, not free text.
      // Exporters that write a company name here are the single most common cause
      // of a Type sheet that looks complete and is not.
      { column: 'Manufacturer', targetSheet: 'Contact', requirement: 'required' },
      { column: 'WarrantyGuarantorParts', targetSheet: 'Contact', requirement: 'required' },
      { column: 'WarrantyGuarantorLabor', targetSheet: 'Contact', requirement: 'required' }
    ]
  },

  Component: {
    name: 'Component',
    requirement: 'required',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'TypeName', requirement: 'required', type: 'text' },
      { name: 'Space', requirement: 'required', type: 'list' },
      { name: 'Description', requirement: 'required', type: 'text' },
      ...EXTERNAL,
      { name: 'SerialNumber', requirement: 'expected', type: 'text' },
      { name: 'InstallationDate', requirement: 'expected', type: 'isoDateTime' },
      { name: 'WarrantyStartDate', requirement: 'expected', type: 'isoDateTime' },
      { name: 'TagNumber', requirement: 'expected', type: 'text' },
      { name: 'BarCode', requirement: 'optional', type: 'text' },
      { name: 'AssetIdentifier', requirement: 'expected', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'TypeName', targetSheet: 'Type', requirement: 'required' },
      // A component can sit in more than one space (a duct run crossing rooms),
      // so this is a list even though the column is singular.
      { column: 'Space', targetSheet: 'Space', list: true, requirement: 'required' }
    ]
  },

  System: {
    name: 'System',
    requirement: 'expected',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'SystemCategory' },
      { name: 'ComponentNames', requirement: 'required', type: 'list' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'ComponentNames', targetSheet: 'Component', list: true, requirement: 'required' }
    ]
  },

  Assembly: {
    name: 'Assembly',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'ParentName', requirement: 'required', type: 'text' },
      { name: 'ChildNames', requirement: 'required', type: 'list' },
      { name: 'AssemblyType', requirement: 'expected', type: 'text' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT],
    dynamicReferences: [
      { sheetColumn: 'SheetName', rowColumn: 'ParentName', requirement: 'required' },
      { sheetColumn: 'SheetName', rowColumn: 'ChildNames', list: true, requirement: 'required' }
    ]
  },

  Connection: {
    name: 'Connection',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'ConnectionType', requirement: 'required', type: 'text' },
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'RowName1', requirement: 'required', type: 'text' },
      { name: 'RowName2', requirement: 'required', type: 'text' },
      { name: 'RealizingElement', requirement: 'optional', type: 'text' },
      { name: 'PortName1', requirement: 'optional', type: 'text' },
      { name: 'PortName2', requirement: 'optional', type: 'text' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT],
    dynamicReferences: [
      { sheetColumn: 'SheetName', rowColumn: 'RowName1', requirement: 'required' },
      { sheetColumn: 'SheetName', rowColumn: 'RowName2', requirement: 'required' }
    ]
  },

  Spare: {
    name: 'Spare',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'SpareCategory' },
      { name: 'TypeName', requirement: 'required', type: 'text' },
      { name: 'Suppliers', requirement: 'required', type: 'list' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'SetNumber', requirement: 'optional', type: 'text' },
      { name: 'PartNumber', requirement: 'optional', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'TypeName', targetSheet: 'Type', requirement: 'required' },
      { column: 'Suppliers', targetSheet: 'Contact', list: true, requirement: 'required' }
    ]
  },

  Resource: {
    name: 'Resource',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'expected', type: 'text' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT]
  },

  Job: {
    name: 'Job',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'JobCategory' },
      { name: 'Status', requirement: 'required', type: 'text', pickList: 'JobStatusType' },
      { name: 'TypeName', requirement: 'required', type: 'text' },
      { name: 'Description', requirement: 'required', type: 'text' },
      { name: 'Duration', requirement: 'required', type: 'number' },
      { name: 'DurationUnit', requirement: 'required', type: 'text', pickList: 'DurationUnit' },
      { name: 'Start', requirement: 'required', type: 'number' },
      { name: 'TaskStartUnit', requirement: 'required', type: 'text', pickList: 'DurationUnit' },
      { name: 'Frequency', requirement: 'required', type: 'number' },
      { name: 'FrequencyUnit', requirement: 'required', type: 'text', pickList: 'DurationUnit' },
      ...EXTERNAL,
      { name: 'TaskNumber', requirement: 'expected', type: 'text' },
      { name: 'Priors', requirement: 'optional', type: 'list' },
      { name: 'ResourceNames', requirement: 'expected', type: 'list' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'TypeName', targetSheet: 'Type', requirement: 'required' },
      { column: 'ResourceNames', targetSheet: 'Resource', list: true, requirement: 'expected' }
    ]
  },

  Impact: {
    name: 'Impact',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'ImpactType', requirement: 'required', type: 'text' },
      { name: 'ImpactStage', requirement: 'required', type: 'text' },
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'RowName', requirement: 'required', type: 'text' },
      { name: 'Value', requirement: 'required', type: 'number' },
      { name: 'ImpactUnit', requirement: 'required', type: 'text' },
      { name: 'LeadInTime', requirement: 'optional', type: 'number' },
      { name: 'Duration', requirement: 'optional', type: 'number' },
      { name: 'LeadOutTime', requirement: 'optional', type: 'number' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' }
    ],
    references: [CREATED_BY_IS_A_CONTACT],
    dynamicReferences: [{ sheetColumn: 'SheetName', rowColumn: 'RowName', requirement: 'required' }]
  },

  Document: {
    name: 'Document',
    requirement: 'expected',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'DocumentCategory' },
      { name: 'ApprovalBy', requirement: 'required', type: 'email' },
      { name: 'Stage', requirement: 'required', type: 'text', pickList: 'StageType' },
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'RowName', requirement: 'required', type: 'text' },
      { name: 'Directory', requirement: 'required', type: 'text' },
      { name: 'File', requirement: 'required', type: 'text' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'Reference', requirement: 'optional', type: 'text' }
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'ApprovalBy', targetSheet: 'Contact', requirement: 'required' }
    ],
    dynamicReferences: [{ sheetColumn: 'SheetName', rowColumn: 'RowName', requirement: 'required' }]
  },

  Attribute: {
    name: 'Attribute',
    requirement: 'expected',
    // Attribute rows are not unique on Name — the whole point is that the same
    // attribute recurs against many rows — so identity is the triple below and
    // the uniqueness rule special-cases it.
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'expected', type: 'text' },
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'RowName', requirement: 'required', type: 'text' },
      { name: 'Value', requirement: 'required', type: 'text' },
      { name: 'Unit', requirement: 'expected', type: 'text' },
      ...EXTERNAL,
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'AllowedValues', requirement: 'optional', type: 'list' }
    ],
    references: [CREATED_BY_IS_A_CONTACT],
    dynamicReferences: [{ sheetColumn: 'SheetName', rowColumn: 'RowName', requirement: 'required' }]
  },

  Coordinate: {
    name: 'Coordinate',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Category', requirement: 'required', type: 'text', pickList: 'CoordinateCategory' },
      { name: 'SheetName', requirement: 'required', type: 'text' },
      { name: 'RowName', requirement: 'required', type: 'text' },
      { name: 'CoordinateXAxis', requirement: 'required', type: 'number' },
      { name: 'CoordinateYAxis', requirement: 'required', type: 'number' },
      { name: 'CoordinateZAxis', requirement: 'required', type: 'number' },
      ...EXTERNAL,
      { name: 'ClockwiseRotation', requirement: 'optional', type: 'number' },
      { name: 'ElevationalRotation', requirement: 'optional', type: 'number' },
      { name: 'YawRotation', requirement: 'optional', type: 'number' }
    ],
    references: [CREATED_BY_IS_A_CONTACT],
    dynamicReferences: [{ sheetColumn: 'SheetName', rowColumn: 'RowName', requirement: 'required' }]
  },

  Issue: {
    name: 'Issue',
    requirement: 'optional',
    key: 'Name',
    columns: [
      { name: 'Name', requirement: 'required', type: 'text' },
      ...CREATED,
      { name: 'Type', requirement: 'required', type: 'text' },
      { name: 'Risk', requirement: 'required', type: 'text' },
      { name: 'Chance', requirement: 'required', type: 'text' },
      { name: 'Impact', requirement: 'required', type: 'text' },
      { name: 'SheetName1', requirement: 'required', type: 'text' },
      { name: 'RowName1', requirement: 'required', type: 'text' },
      { name: 'SheetName2', requirement: 'optional', type: 'text' },
      { name: 'RowName2', requirement: 'optional', type: 'text' },
      { name: 'Description', requirement: 'expected', type: 'text' },
      { name: 'Owner', requirement: 'expected', type: 'email' },
      { name: 'Mitigation', requirement: 'optional', type: 'text' },
      ...EXTERNAL
    ],
    references: [
      CREATED_BY_IS_A_CONTACT,
      { column: 'Owner', targetSheet: 'Contact', requirement: 'expected' }
    ],
    dynamicReferences: [
      { sheetColumn: 'SheetName1', rowColumn: 'RowName1', requirement: 'required' },
      { sheetColumn: 'SheetName2', rowColumn: 'RowName2', requirement: 'optional' }
    ]
  }
};

/** The sheet holding a COBie file's own enumerations, if it ships one. */
export const PICKLIST_SHEET = 'PickLists';

/**
 * Fall-back enumerations for the pick lists whose values COBie fixes rather than
 * leaving to the project. Used only when the file has no PickLists sheet of its
 * own — a file that ships one is checked against its own lists, because a
 * project is entitled to extend them.
 */
export const BUILT_IN_PICK_LISTS: Readonly<Record<string, readonly string[]>> = {
  AssetType: ['Fixed', 'Moveable'],
  DurationUnit: [
    'year', 'month', 'week', 'day', 'hour', 'minute', 'second'
  ],
  StageType: [
    'Requirement', 'Design', 'Construction', 'Handover', 'Operations',
    'As-Built', 'Commissioning'
  ],
  JobStatusType: ['Required', 'Optional', 'Recommended', 'Not Required']
};

export function sheetDefinition(name: CobieSheetName): CobieSheet {
  return COBIE_SCHEMA[name];
}

/** True when `name` is one of the sheets the schema defines. */
export function isCobieSheet(name: string): name is CobieSheetName {
  return Object.prototype.hasOwnProperty.call(COBIE_SCHEMA, name);
}

/**
 * Every header spelling that resolves to `column`, lowercased for matching.
 * Header comparison is case- and space-insensitive throughout: exporters vary on
 * "ExtSystem" versus "Ext System", and rejecting those would report thousands of
 * spurious defects on files that are otherwise fine.
 */
export function headerCandidates(column: CobieColumn): string[] {
  const names = [column.name].concat(column.aliases ? Array.from(column.aliases) : []);
  return names.map(normaliseHeader);
}

export function normaliseHeader(header: string): string {
  return header.replace(/\s+/g, '').toLowerCase();
}
