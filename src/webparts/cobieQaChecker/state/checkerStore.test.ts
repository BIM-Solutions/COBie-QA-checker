import { applyFilters } from './checkerStore';
import type { Finding } from '../models/findings';

const FINDINGS: Finding[] = [
  { ruleId: 'completeness.missingValue', category: 'completeness', severity: 'error', sheet: 'Space', row: 2, column: 'Category', message: 'Category is empty.' },
  { ruleId: 'reference.broken', category: 'reference', severity: 'error', sheet: 'Component', row: 5, column: 'TypeName', value: 'AHU-99', message: '"AHU-99" is not a Type Name.' },
  { ruleId: 'pickList.value', category: 'pickList', severity: 'warning', sheet: 'Type', row: 3, column: 'AssetType', value: 'Portable', message: '"Portable" is not in the AssetType pick list.' }
];

describe('applyFilters', () => {
  it('returns everything when no filter is set', () => {
    // Empty arrays mean "no filter", not "match nothing". Getting this
    // backwards shows an empty table on a file with thousands of findings.
    expect(applyFilters(FINDINGS, { severities: [], categories: [], sheets: [], search: '' }))
      .toHaveLength(3);
  });

  it('filters by severity', () => {
    const result = applyFilters(FINDINGS, { severities: ['warning'], categories: [], sheets: [], search: '' });
    expect(result.map((f) => f.ruleId)).toEqual(['pickList.value']);
  });

  it('filters by category', () => {
    const result = applyFilters(FINDINGS, { severities: [], categories: ['reference'], sheets: [], search: '' });
    expect(result.map((f) => f.ruleId)).toEqual(['reference.broken']);
  });

  it('filters by sheet', () => {
    const result = applyFilters(FINDINGS, { severities: [], categories: [], sheets: ['Space'], search: '' });
    expect(result.map((f) => f.ruleId)).toEqual(['completeness.missingValue']);
  });

  it('combines filters as AND', () => {
    const result = applyFilters(FINDINGS, {
      severities: ['error'], categories: ['reference'], sheets: [], search: ''
    });
    expect(result.map((f) => f.ruleId)).toEqual(['reference.broken']);
  });

  it('searches the value column', () => {
    const result = applyFilters(FINDINGS, { severities: [], categories: [], sheets: [], search: 'ahu-99' });
    expect(result.map((f) => f.ruleId)).toEqual(['reference.broken']);
  });

  it('searches case-insensitively', () => {
    const result = applyFilters(FINDINGS, { severities: [], categories: [], sheets: [], search: 'PORTABLE' });
    expect(result.map((f) => f.ruleId)).toEqual(['pickList.value']);
  });

  it('does not match on the rule id', () => {
    // "missingValue" appears only in a rule id, never in a column the table
    // renders. Matching it would return a row for a term the user cannot see.
    expect(applyFilters(FINDINGS, { severities: [], categories: [], sheets: [], search: 'missingvalue' }))
      .toHaveLength(0);
  });

  it('ignores surrounding whitespace in the search term', () => {
    const result = applyFilters(FINDINGS, { severities: [], categories: [], sheets: [], search: '  ahu-99  ' });
    expect(result).toHaveLength(1);
  });
});
