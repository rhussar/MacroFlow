export const INITIAL_SEARCH_DATA = {
  status: 'idle',
  workbook: null,
  modules: [],
  macros: [],
  shortcutAudit: null,
  error: null
};

export const SEARCH_FOCUS_REFRESH_COOLDOWN_MS = 600;
export const SEARCH_FULL_REFRESH_STALE_MS = 12000;
export const SEARCH_PERIODIC_REFRESH_MS = 30000;
export const SEARCH_HELPER_EVENT_DEBOUNCE_MS = 150;
export const SEARCH_HELPER_EVENT_COOLDOWN_MS = 1500;
export const SHORTCUT_REFRESH_TTL_MS = 5000;
export const PERSONAL_REFRESH_TTL_MS = 30000;
