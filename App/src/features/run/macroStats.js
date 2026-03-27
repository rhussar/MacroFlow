const STORAGE_KEY = 'macroflow-stats';

const DEFAULT_STATS = { totalRuns: 0, totalDurationMs: 0, totalTimeSavedMs: 0 };
const TIME_SAVED_MULTIPLIER = 9;

export function getMacroStats() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_STATS };
    const parsed = JSON.parse(raw);
    return {
      totalRuns: Number(parsed.totalRuns) || 0,
      totalDurationMs: Number(parsed.totalDurationMs) || 0,
      totalTimeSavedMs: Number(parsed.totalTimeSavedMs) || 0,
    };
  } catch {
    return { ...DEFAULT_STATS };
  }
}

export function recordMacroRun(durationMs) {
  try {
    const stats = getMacroStats();
    stats.totalRuns += 1;
    stats.totalDurationMs += durationMs;
    stats.totalTimeSavedMs += durationMs * TIME_SAVED_MULTIPLIER;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(stats));
  } catch {
    // Silent fail if localStorage is unavailable.
  }
}

export function formatTimeSaved(ms) {
  if (ms < 1000) return '0 sec';
  if (ms < 60_000) return `${Math.round(ms / 1000)} sec`;
  if (ms < 3_600_000) return `${(ms / 60_000).toFixed(1)} min`;
  return `${(ms / 3_600_000).toFixed(1)} hrs`;
}
