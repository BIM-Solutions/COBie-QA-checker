# COBie QA Checker

A SharePoint Framework web part that checks COBie deliverables against the
COBie 2.4 schema, without the file leaving Microsoft 365.

The deliverable is already in a SharePoint document library. Downloading it,
checking it in a desktop tool and uploading a report is the workflow this
replaces: pick the file from the library, read the findings, export the report
back to the library.

## What it checks

Seven families of rule, all driven from one schema definition
(`models/cobieSchema.ts`) rather than hard-coded:

| Kind | What it catches |
| --- | --- |
| **Structure** | Missing or empty sheets; missing columns; sheets the schema does not recognise |
| **Missing values** | Blank cells in columns COBie requires |
| **Placeholder values** | `n/a`, `TBC`, `-`, `Unknown` and friends standing in for a value |
| **Duplicates** | Repeated `Name` (or `Email` on Contact) within a sheet |
| **Broken references** | `Space→Floor`, `Component→Type`, `Zone→Space`, `System→Component`, `Type.Manufacturer→Contact`, and the polymorphic `SheetName`+`RowName` pairs used by Attribute, Document, Coordinate, Impact, Assembly, Connection and Issue |
| **Format** | ISO 8601 dates (including dates that parse but do not exist), email addresses, numerics |
| **Pick lists** | Values outside their enumeration, checked against the file's own `PickLists` sheet where it has one |

Severity follows the schema's requirement level and nothing else: `required`
produces an error, `expected` a warning, `optional` is not reported. A file
passes when it has no errors.

### Two numbers worth understanding

**Completeness** is the share of *required* cells holding a real value —
neither blank nor a placeholder. It is weighted per cell, not per sheet, so a
complete four-row Facility sheet cannot offset a threadbare 40,000-row
Component sheet. A required column the file omits counts as wholly unfilled,
because otherwise a file would score higher for dropping a column than for
including it and leaving it blank.

**The findings list is capped at 5,000**; the error and warning counts are not.
A file missing its Contact sheet generates a finding per row per email column,
which is hundreds of thousands on a real deliverable. The counts stay exact so
the summary never understates the problem.

## Reading COBie files

Accepts `.xlsx`, `.xlsm`, `.csv`, `.tsv` and `.txt`. A CSV holds one sheet and
its file name is the sheet name — `Component.csv` is the Component sheet.

Header matching is case- and space-insensitive, and accepts the alternate
spellings real exporters emit (`ExternalSystem` for `ExtSystem`). A header row
sitting below a title band is found rather than assumed to be row 1. Every
finding carries the real spreadsheet row number, so it can be looked up in
Excel directly.

## Where things are stored

Everything stays inside Microsoft 365 — no Azure, no database, no third-party
service.

- **COBie files** are read from the document library set in the property pane
  (default: `Shared Documents`).
- **Reports** are exported as a two-sheet Excel workbook — a summary and a
  findings work list — written back to the library.
- **Run history** is one row per check in a `COBie Check History` list, which
  the web part creates on the site the first time it needs it. Only the
  summary is stored, never the findings: a single run would blow past the
  5,000-item list view threshold on its own. Recording is best-effort, so a
  user with read-only access to the site can still check files.

## Configuration

Three property-pane settings, all optional:

| Setting | Default |
| --- | --- |
| COBie files library | `Shared Documents` on this site |
| Reports library | Same library as the file being checked |
| Record each check to this site | On |

Library paths accept `Shared Documents`, `/sites/Project/Shared Documents/COBie`
or a full absolute URL.

## Building

Requires Node 22 (SPFx 1.23.x supports Node 22 LTS only).

```bash
npm ci
npm test        # heft test - build, lint and tests; the whole gate
npm start       # local workbench
npm run build   # production build and .sppkg
```

CI runs `npx heft test --clean` plus `npx eslint src --max-warnings 0` on every
pull request. Both must pass.

The package deploys as `solution/cobie-qa-checker.sppkg`. It is tenant-scoped
with `skipFeatureDeployment`, so an admin uploads it to the app catalogue once
and any site can add the web part.

## Conventions the code depends on

- **`target: es5` with no `downlevelIteration`.** Use `Array.from` and
  `.forEach` rather than spreading or iterating Sets and Maps directly.
- **Explicit return types on exported functions.** The lint rule is a warning
  and CI treats it as an error, which is why it runs eslint separately with
  `--max-warnings 0`.
- **The xlsx reader and writer are lazy chunks** (~300KB and ~250KB). A page
  that merely hosts the checker should not pay for them. Import them
  dynamically; a static import silently moves them into the entry bundle.
- **Severity mapping lives in one function** (`severityFor` in `rules.ts`), so
  "what counts as a failure" is one decision rather than thirty.
- **Blank and placeholder are different findings.** A blank cell is usually an
  unconfigured exporter; an `n/a` is a person who declined to fill the field.
  Different owners, different fixes — the report must not merge them.

## What this does not do

- **No IFC.** It checks the COBie spreadsheet, not the model it came from.
- **No classification validation.** `Category` is checked against the file's
  own `PickLists` sheet when it has one, and otherwise not at all —
  Uniclass and OmniClass are project decisions, and asserting one would report
  a defect on every correctly classified file.
- **No custom rulesets.** The schema is COBie 2.4 as published. Narrowing it
  to a client's own MIDP is a change to `cobieSchema.ts`, which is where that
  seam deliberately sits.
- **No findings in run history.** By design; see above.
