import { runChecks } from './RuleEngine';
import { CONTACT, cleanWorkbook, sheet, workbook } from './fixtures';

const OPTIONS = { checkedBy: 'tester' };

describe('runChecks', () => {
  it('passes a file with no error-severity findings', () => {
    // The suite's anchor. Without it, a rule that fired on everything would
    // still satisfy every "reports X" test in the codebase.
    const run = runChecks(cleanWorkbook(), OPTIONS);
    expect(run.errorCount).toBe(0);
    expect(run.passed).toBe(true);
  });

  it('still warns about the expected sheets a clean file omits', () => {
    const run = runChecks(cleanWorkbook(), OPTIONS);
    expect(run.warningCount).toBeGreaterThan(0);
    expect(run.passed).toBe(true);
  });

  it('fails a file with any error', () => {
    const run = runChecks(workbook([CONTACT]), OPTIONS);
    expect(run.passed).toBe(false);
    expect(run.errorCount).toBeGreaterThan(0);
  });

  it('sorts errors before warnings', () => {
    const run = runChecks(workbook([CONTACT]), OPTIONS);
    const firstWarning = run.findings.map((f) => f.severity).indexOf('warning');
    const lastError = run.findings.map((f) => f.severity).lastIndexOf('error');
    expect(lastError).toBeLessThan(firstWarning);
  });

  it('turns a read warning into an error finding', () => {
    // A sheet that failed to decode would otherwise score zero with nothing to
    // say the checker never read it.
    const wb = {
      fileName: 'broken.xlsx',
      sheets: cleanWorkbook().sheets,
      readWarnings: ['Sheet "Component" could not be read: corrupt']
    };
    const run = runChecks(wb, OPTIONS);
    const finding = run.findings.filter((f) => f.ruleId === 'structure.unreadable')[0];
    expect(finding.severity).toBe('error');
    expect(finding.message).toContain('corrupt');
  });

  it('caps the findings it keeps but not the counts it reports', () => {
    const rows: string[][] = [['Name', 'CreatedBy', 'CreatedOn', 'Category']];
    for (let i = 0; i < 200; i++) { rows.push([`L${i}`, '', '', '']); }
    const run = runChecks(workbook([sheet('Floor', rows)]), { checkedBy: 't', maxFindings: 10 });

    expect(run.findings).toHaveLength(10);
    // The counts must stay honest: a truncated list that also truncated the
    // totals would tell the user their file has ten problems when it has
    // hundreds.
    expect(run.errorCount).toBeGreaterThan(10);
  });

  it('records the file name and who ran the check', () => {
    const run = runChecks(cleanWorkbook(), { checkedBy: 'ann@example.com' });
    expect(run.fileName).toBe('test.xlsx');
    expect(run.checkedBy).toBe('ann@example.com');
    expect(run.checkedOn).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  });
});

describe('completeness', () => {
  it('scores a fully populated file at 1', () => {
    expect(runChecks(cleanWorkbook(), OPTIONS).completeness).toBe(1);
  });

  it('scores an empty file at 0 rather than dividing by zero', () => {
    // `0/0` would otherwise report a perfect score for a workbook with nothing
    // in it - the single most flattering possible lie.
    expect(runChecks(workbook([]), OPTIONS).completeness).toBe(0);
  });

  it('counts a placeholder as incomplete', () => {
    const wb = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', 'n/a']
    ])]);
    // Four required columns, one placeholder.
    expect(runChecks(wb, OPTIONS).completeness).toBeCloseTo(0.75, 5);
  });

  it('counts a missing required column as wholly unfilled', () => {
    // Otherwise a file scores *higher* for omitting a column than for
    // including it and leaving it blank, which is exactly backwards.
    const withColumn = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', '']
    ])]);
    const withoutColumn = workbook([sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00']
    ])]);
    expect(runChecks(withoutColumn, OPTIONS).completeness)
      .toBeCloseTo(runChecks(withColumn, OPTIONS).completeness, 5);
  });

  it('weights by cell, so a big sparse sheet is not offset by a small full one', () => {
    // Averaging the per-sheet percentages would let a complete four-row
    // Facility sheet flatter a threadbare 40,000-row Component sheet.
    const rows: string[][] = [['Name', 'CreatedBy', 'CreatedOn', 'Category']];
    for (let i = 0; i < 100; i++) { rows.push([`L${i}`, '', '', '']); }

    const wb = workbook([
      sheet('Facility', [
        ['Name', 'CreatedBy', 'CreatedOn', 'Category', 'ProjectName', 'SiteName',
          'LinearUnits', 'AreaUnits', 'VolumeUnits', 'CurrencyUnit'],
        ['HQ', 'a@b.com', '2026-01-15T09:30:00', 'Office', 'P', 'S', 'mm', 'm2', 'm3', 'GBP']
      ]),
      sheet('Floor', rows)
    ]);

    // Floor contributes 400 required cells of which 100 are filled; Facility
    // contributes 10 of 10. A sheet-averaged score would be ~0.63.
    expect(runChecks(wb, OPTIONS).completeness).toBeLessThan(0.3);
  });
});

describe('sheet summaries', () => {
  it('lists every schema sheet, present or not', () => {
    const run = runChecks(cleanWorkbook(), OPTIONS);
    const zone = run.sheets.filter((s) => s.sheet === 'Zone')[0];
    expect(zone.present).toBe(false);
    expect(zone.rowCount).toBe(0);
    // Undefined rather than 0: no measurement was taken, and rendering 0%
    // would claim one was.
    expect(zone.completeness).toBeUndefined();
  });

  it('counts errors and warnings against the sheet they belong to', () => {
    const wb = workbook([CONTACT, sheet('Floor', [
      ['Name', 'CreatedBy', 'CreatedOn', 'Category'],
      ['L0', 'a@b.com', '2026-01-15T09:30:00', '']
    ])]);
    const floor = runChecks(wb, OPTIONS).sheets.filter((s) => s.sheet === 'Floor')[0];
    expect(floor.present).toBe(true);
    expect(floor.rowCount).toBe(1);
    expect(floor.errors).toBeGreaterThan(0);
  });
});
