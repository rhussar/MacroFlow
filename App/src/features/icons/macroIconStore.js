/**
 * macroIconStore.js
 *
 * Persists icon assignments per macro to localStorage.
 * Each macro can have one ImageMSO icon name assigned.
 */

const STORAGE_KEY = 'macroflow-icon-assignments';

let cache = null;

/**
 * Normalize a macro ID to just workbook::macro::module::name,
 * stripping the kind::scope suffix (e.g. Sub::Public, Sub::Implicit)
 * so icon assignments match regardless of how the ID was constructed.
 */
function normalizeKey(macroId) {
  const id = String(macroId || '');
  const markerIndex = id.indexOf('::macro::');
  if (markerIndex >= 0) {
    // Format: workbookKey::macro::module::name::kind::scope
    // Keep: workbookKey::macro::module::name
    const afterMarker = id.slice(markerIndex + 9); // after "::macro::"
    const parts = afterMarker.split('::');
    // module::name are the first two parts
    if (parts.length >= 2) {
      return `${id.slice(0, markerIndex)}::macro::${parts[0]}::${parts[1]}`;
    }
  }
  // Format without ::macro:: marker: module::name::kind::scope
  const parts = id.split('::');
  if (parts.length >= 2) {
    return `${parts[0]}::${parts[1]}`;
  }
  return id;
}

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

/**
 * Extract the short module::name key (without workbook prefix) from any macro ID format.
 */
function toShortKey(macroId) {
  const id = String(macroId || '');
  const markerIndex = id.indexOf('::macro::');
  if (markerIndex >= 0) {
    return normalizeKey(id.slice(markerIndex + 9));
  }
  return normalizeKey(id);
}

export function getMacroIcon(macroId) {
  const data = load();
  const key = normalizeKey(macroId);
  if (data[key]) return data[key];
  if (data[macroId]) return data[macroId];

  // Cross-format fallback: active-workbook macros store icons under
  // "module::name" while Create tab stores "workbookKey::macro::module::name"
  // (and vice-versa). Try the short key so both formats find each other.
  const shortKey = toShortKey(macroId);
  if (shortKey !== key && data[shortKey]) return data[shortKey];

  // Reverse: lookup key is short but stored key has a workbook prefix.
  // Scan stored keys for a matching short form.
  for (const storedKey of Object.keys(data)) {
    if (toShortKey(storedKey) === shortKey) return data[storedKey];
  }

  return null;
}

function notifyChange() {
  window.dispatchEvent(new CustomEvent('macroflow-icon-change'));
}

export function setMacroIcon(macroId, iconName) {
  load();
  const key = normalizeKey(macroId);
  const shortKey = toShortKey(macroId);

  // Remove any stale entry stored under a different format for the same macro
  // so we don't accumulate both "module::name" and "wbKey::macro::module::name".
  if (shortKey !== key) {
    delete cache[shortKey];
  }
  for (const storedKey of Object.keys(cache)) {
    if (storedKey !== key && toShortKey(storedKey) === shortKey) {
      delete cache[storedKey];
    }
  }

  cache[key] = iconName;
  save();
  notifyChange();
}

export function removeMacroIcon(macroId) {
  load();
  const key = normalizeKey(macroId);
  const shortKey = toShortKey(macroId);
  delete cache[key];
  // Also remove any cross-format entries for the same macro
  if (shortKey !== key) {
    delete cache[shortKey];
  }
  for (const storedKey of Object.keys(cache)) {
    if (storedKey !== key && toShortKey(storedKey) === shortKey) {
      delete cache[storedKey];
    }
  }
  save();
  notifyChange();
}

export function renameMacroIcon(oldMacroId, newMacroId) {
  if (!oldMacroId || !newMacroId || oldMacroId === newMacroId) return;
  const icon = getMacroIcon(oldMacroId);
  if (!icon) return;
  removeMacroIcon(oldMacroId);
  setMacroIcon(newMacroId, icon);
}

export function getAllMacroIcons() {
  return { ...load() };
}
