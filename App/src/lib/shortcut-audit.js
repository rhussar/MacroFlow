function toSafeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

export function normalizeShortcutKey(input) {
  const raw = toSafeString(input).toUpperCase();
  if (!raw) {
    return '';
  }

  const collapsed = raw.replace(/\s+/g, ' ');
  const tokens = collapsed
    .split('+')
    .map((token) => token.trim())
    .filter(Boolean);

  if (tokens.length === 0) {
    return '';
  }

  if (tokens.length === 1) {
    return tokens[0];
  }

  return tokens.join('+');
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
