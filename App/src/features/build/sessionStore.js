/**
 * sessionStore.js
 *
 * Persists AI chat sessions to localStorage so they survive app restarts.
 * Sessions are capped at MAX_SESSIONS (FIFO) to stay within quota.
 */

const STORAGE_KEY = 'macroflow-sessions';
const MAX_SESSIONS = 30;

let cache = null;

function load() {
  if (cache) return cache;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    cache = Array.isArray(parsed) ? parsed : [];
  } catch {
    cache = [];
  }
  return cache;
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cache));
  } catch {
    // Quota exceeded — prune oldest 5 and retry once
    try {
      cache = cache.slice(cache.length > 5 ? 5 : 0);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cache));
    } catch { /* still failed — silently ignore */ }
  }
}

export function loadSessions() {
  return [...load()];
}

export function saveSessions(sessions) {
  cache = Array.isArray(sessions) ? sessions : [];
  if (cache.length > MAX_SESSIONS) {
    cache = cache.slice(cache.length - MAX_SESSIONS);
  }
  save();
}
