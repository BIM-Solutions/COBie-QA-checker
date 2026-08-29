import { create } from 'zustand';
import type { CheckRun, FindingCategory, Severity } from '../models/findings';

/**
 * The one piece of state the screens share: the current run and how the issues
 * table is filtered.
 *
 * A store rather than lifted state because the filters are set on the issues
 * screen, summarised in the header and cleared from the dashboard's category
 * tiles - three places, none of which owns the others.
 */

export interface Filters {
  readonly severities: readonly Severity[];
  readonly categories: readonly FindingCategory[];
  readonly sheets: readonly string[];
  readonly search: string;
}

export type Screen = 'source' | 'results';

export interface CheckerState {
  readonly run?: CheckRun;
  readonly screen: Screen;
  readonly busy: boolean;
  /** What the app is currently doing, shown next to the spinner. */
  readonly busyMessage: string;
  readonly error?: string;
  /** Set when the check succeeded but recording it to the list did not. */
  readonly notice?: string;
  readonly filters: Filters;

  setRun(run: CheckRun): void;
  setBusy(busy: boolean, message?: string): void;
  setError(error: string | undefined): void;
  setNotice(notice: string | undefined): void;
  setScreen(screen: Screen): void;
  setFilters(filters: Partial<Filters>): void;
  clearFilters(): void;
  reset(): void;
}

/** Empty arrays mean "no filter", not "match nothing". */
const NO_FILTERS: Filters = { severities: [], categories: [], sheets: [], search: '' };

export const useCheckerStore = create<CheckerState>((set) => ({
  screen: 'source',
  busy: false,
  busyMessage: '',
  filters: NO_FILTERS,

  setRun: (run) => set({
    run,
    screen: 'results',
    // A new run must not inherit the previous one's filters: the sheets they
    // name may not exist in this file, and the user would see an empty table
    // with no obvious cause.
    filters: NO_FILTERS,
    error: undefined
  }),

  setBusy: (busy, message) => set({ busy, busyMessage: busy ? (message || '') : '' }),
  setError: (error) => set({ error, busy: false, busyMessage: '' }),
  setNotice: (notice) => set({ notice }),
  setScreen: (screen) => set({ screen }),
  setFilters: (filters) => set((state) => ({ filters: { ...state.filters, ...filters } })),
  clearFilters: () => set({ filters: NO_FILTERS }),
  reset: () => set({
    run: undefined,
    screen: 'source',
    busy: false,
    busyMessage: '',
    error: undefined,
    notice: undefined,
    filters: NO_FILTERS
  })
}));

/**
 * Applies the filters. Kept out of the store and out of the component so it can
 * be tested directly - it is the one place where a wrong answer looks entirely
 * plausible on screen.
 */
export function applyFilters<T extends {
  severity: Severity;
  category: FindingCategory;
  sheet: string;
  message: string;
  column?: string;
  value?: string;
}>(findings: readonly T[], filters: Filters): T[] {
  const search = filters.search.trim().toLowerCase();

  return findings.filter((finding) => {
    if (filters.severities.length > 0 && filters.severities.indexOf(finding.severity) === -1) {
      return false;
    }
    if (filters.categories.length > 0 && filters.categories.indexOf(finding.category) === -1) {
      return false;
    }
    if (filters.sheets.length > 0 && filters.sheets.indexOf(finding.sheet) === -1) {
      return false;
    }
    if (search === '') { return true; }

    // Searches the columns a user can see. Matching the rule id as well would
    // surface rows for a term that appears nowhere on screen.
    const haystack = [finding.message, finding.column || '', finding.value || '', finding.sheet]
      .join(' ')
      .toLowerCase();
    return haystack.indexOf(search) !== -1;
  });
}
