import { parseShortcutLetter } from './shortcut-keybind.js';

function toSafeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function parseShortcutParts(input) {
  const raw = toSafeString(input);
  if (!raw) {
    return null;
  }

  const collapsed = raw.replace(/\s+/g, ' ');
  const tokens = collapsed
    .split('+')
    .map((token) => token.trim())
    .filter(Boolean);

  if (tokens.length === 0) {
    return null;
  }

  const modifierTokens = tokens.length > 1 ? tokens.slice(0, -1) : [];
  const keyToken = tokens[tokens.length - 1];
  if (!keyToken) {
    return null;
  }

  let shift = false;
  let alt = false;
  modifierTokens.forEach((token) => {
    const normalized = token.toLowerCase();
    if (normalized === 'shift') {
      shift = true;
    } else if (normalized === 'alt') {
      alt = true;
    }
  });

  const isLetter = /^[a-zA-Z]$/.test(keyToken);
  const key = keyToken.toUpperCase();

  // Infer Shift from letter casing (A vs a) for both compact and tokenized inputs.
  if (isLetter && keyToken === keyToken.toUpperCase() && keyToken !== keyToken.toLowerCase()) {
    shift = true;
  }

  return {
    // Excel macro shortcuts are always Ctrl-based; Shift is encoded by letter case.
    ctrl: true,
    shift,
    alt,
    key,
    isLetter
  };
}

export function normalizeShortcutKey(input) {
  const parts = parseShortcutParts(input);
  if (!parts) {
    return '';
  }

  const modifierParts = ['Ctrl'];
  if (parts.shift) {
    modifierParts.push('Shift');
  }
  if (parts.alt) {
    modifierParts.push('Alt');
  }

  const displayKey = parts.isLetter
    ? (parts.shift ? parts.key : parts.key.toLowerCase())
    : parts.key;

  return `${modifierParts.join(' + ')} + ${displayKey}`;
}

export function toExcelShortcutKey(input) {
  const parts = parseShortcutParts(input);
  if (!parts) {
    return '';
  }

  if (parts.isLetter) {
    return parts.shift ? parts.key.toUpperCase() : parts.key.toLowerCase();
  }

  return parts.key;
}

export function canonicalizeMacroIdentity(value) {
  const normalized = toSafeString(value);
  if (!normalized) {
    return '';
  }

  const bangIndex = normalized.lastIndexOf('!');
  const withoutWorkbook = bangIndex >= 0 && bangIndex < normalized.length - 1
    ? normalized.slice(bangIndex + 1)
    : normalized;

  return withoutWorkbook.trim().toLowerCase();
}

export function toCompactMacroName(fullName) {
  const normalized = toSafeString(fullName);
  if (!normalized) {
    return '';
  }

  const bangIndex = normalized.lastIndexOf('!');
  if (bangIndex >= 0 && bangIndex < normalized.length - 1) {
    return normalized.slice(bangIndex + 1).trim();
  }

  return normalized;
}

function normalizeMappedRow(row) {
  const macroFull = toSafeString(row?.macro || row?.macroName || row?.fullName);
  const shortcutRaw = toSafeString(row?.shortcut || row?.shortcutKey || row?.key);
  if (!macroFull || !shortcutRaw) {
    return null;
  }

  const shortcutNorm = normalizeShortcutKey(shortcutRaw);
  if (!shortcutNorm) {
    return null;
  }

  return {
    macroFull,
    macroCompact: toCompactMacroName(macroFull),
    shortcutRaw,
    shortcutNorm
  };
}

function normalizeUnmappedRow(row) {
  const macroFull = typeof row === 'string'
    ? toSafeString(row)
    : toSafeString(row?.macro || row?.macroName || row?.fullName);

  if (!macroFull) {
    return null;
  }

  return {
    macroFull,
    macroCompact: toCompactMacroName(macroFull)
  };
}

function sortByMacro(a, b) {
  const aKey = `${a.macroCompact}::${a.macroFull}`.toLowerCase();
  const bKey = `${b.macroCompact}::${b.macroFull}`.toLowerCase();
  return aKey.localeCompare(bKey);
}

export function normalizeAuditResponse(apiResult) {
  const mappedSource = Array.isArray(apiResult?.shortcuts) ? apiResult.shortcuts : [];
  const unmappedSource = Array.isArray(apiResult?.unmapped) ? apiResult.unmapped : [];

  const mapped = mappedSource
    .map((row) => normalizeMappedRow(row))
    .filter(Boolean)
    .sort((a, b) => {
      const shortcutOrder = a.shortcutNorm.localeCompare(b.shortcutNorm);
      if (shortcutOrder !== 0) {
        return shortcutOrder;
      }
      return sortByMacro(a, b);
    });

  const unmapped = unmappedSource
    .map((row) => normalizeUnmappedRow(row))
    .filter(Boolean)
    .sort(sortByMacro);

  const conflictMap = new Map();
  mapped.forEach((row) => {
    const rows = conflictMap.get(row.shortcutNorm) || [];
    rows.push(row);
    conflictMap.set(row.shortcutNorm, rows);
  });

  const conflicts = [];
  conflictMap.forEach((rows, shortcutNorm) => {
    const uniqueByMacro = new Map();
    rows.forEach((row) => {
      const uniqueKey = row.macroFull.toLowerCase();
      if (!uniqueByMacro.has(uniqueKey)) {
        uniqueByMacro.set(uniqueKey, {
          macroFull: row.macroFull,
          macroCompact: row.macroCompact
        });
      }
    });

    if (uniqueByMacro.size <= 1) {
      return;
    }

    const entries = Array.from(uniqueByMacro.values()).sort(sortByMacro);
    conflicts.push({
      shortcutNorm,
      entries
    });
  });

  conflicts.sort((a, b) => a.shortcutNorm.localeCompare(b.shortcutNorm));

  return {
    mapped,
    unmapped,
    conflicts,
    note: toSafeString(apiResult?.note)
  };
}

export function mapAuditShortcutsToMacroIds(apiResult, macros = []) {
  const mappedRows = normalizeAuditResponse(apiResult).mapped;
  const macroKeyToId = buildMacroKeyToId(macros);

  const shortcutByMacroId = {};
  mappedRows.forEach((row) => {
    const canonical = canonicalizeMacroIdentity(row.macroFull);
    const macroId = macroKeyToId.get(canonical);
    if (!macroId) {
      return;
    }
    shortcutByMacroId[macroId] = row.shortcutNorm;
  });

  return shortcutByMacroId;
}

function buildMacroKeyToId(macros = []) {
  const sourceMacros = Array.isArray(macros) ? macros : [];
  const macroKeyToId = new Map();

  sourceMacros.forEach((macro) => {
    const identity = toSafeString(macro?.fullName || macro?.runTarget || '');
    const fallbackIdentity =
      identity || `${toSafeString(macro?.module)}.${toSafeString(macro?.name)}`;
    const canonical = canonicalizeMacroIdentity(fallbackIdentity);
    if (canonical && !macroKeyToId.has(canonical) && macro?.id) {
      macroKeyToId.set(canonical, macro.id);
    }
  });

  return macroKeyToId;
}

export function mapAuditConflictsByShortcut(apiResult, macros = []) {
  const mappedRows = normalizeAuditResponse(apiResult).mapped;
  const macroKeyToId = buildMacroKeyToId(macros);
  const shortcutGroups = new Map();

  mappedRows.forEach((row) => {
    const shortcutLetter = parseShortcutLetter(row.shortcutNorm);
    if (!shortcutLetter) {
      return;
    }

    const canonical = canonicalizeMacroIdentity(row.macroFull);
    const macroId = macroKeyToId.get(canonical);
    if (!macroId) {
      return;
    }

    if (!shortcutGroups.has(shortcutLetter)) {
      shortcutGroups.set(shortcutLetter, new Set());
    }

    shortcutGroups.get(shortcutLetter).add(macroId);
  });

  const conflictsByShortcut = {};
  Array.from(shortcutGroups.entries())
    .sort(([shortcutA], [shortcutB]) => shortcutA.localeCompare(shortcutB))
    .forEach(([shortcutLetter, macroIdSet]) => {
      if (macroIdSet.size <= 1) {
        return;
      }

      conflictsByShortcut[shortcutLetter] = Array.from(macroIdSet).sort();
    });

  return conflictsByShortcut;
}
