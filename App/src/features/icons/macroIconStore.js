/**
 * macroIconStore.js
 *
 * Persists icon assignments per macro to localStorage.
 * Each macro can have one ImageMSO icon name assigned.
 */

const STORAGE_KEY = 'macroflow-icon-assignments';

let cache = null;

function load() {
  if (cache) return cache;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    cache = raw ? JSON.parse(raw) : {};
  } catch {
    cache = {};
  }
  return cache;
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cache));
  } catch { /* quota exceeded — silently ignore */ }
}

export function getMacroIcon(macroId) {
  return load()[macroId] || null;
}

export function setMacroIcon(macroId, iconName) {
  load();
  cache[macroId] = iconName;
  save();
}

export function removeMacroIcon(macroId) {
  load();
  delete cache[macroId];
  save();
}

export function getAllMacroIcons() {
  return { ...load() };
}
