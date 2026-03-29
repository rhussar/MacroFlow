import { normalizeModules } from '../../lib/search-data.js';

function toSafeString(value) {
  return String(value || '').trim();
}

export function getWorkbookKey(workbook) {
  const path = toSafeString(workbook?.path);
  const name = toSafeString(workbook?.name);
  return path || name || '';
}

export function toWorkbookRequest(workbook) {
  return {
    workbookName: toSafeString(workbook?.name),
    workbookPath: toSafeString(workbook?.path)
  };
}

export function toWorkbookModel(workbook, options = {}) {
  const fallback = options?.fallback || null;
  const defaultName = toSafeString(options?.defaultName) || 'Workbook';
  const name = toSafeString(workbook?.name || fallback?.name);
  const path = toSafeString(workbook?.path || fallback?.path);
  const key = toSafeString(workbook?.key || fallback?.key || path || name);

  if (!name && !path && !key) {
    return null;
  }

  return {
    name: name || defaultName,
    path,
    key: key || path || name
  };
}

export function normalizeListContextModules(rawModules = []) {
  const source = Array.isArray(rawModules) ? rawModules : [];
  const normalized = [];

  source.forEach((moduleItem) => {
    const workbook = {
      name: toSafeString(moduleItem?.workbookName),
      path: toSafeString(moduleItem?.workbookPath)
    };
    const rows = normalizeModules([moduleItem], workbook);
    rows.forEach((row) => normalized.push(row));
  });

  return normalized;
}

export function sortWorkbooksForPicker(workbooks, activeWorkbookKey) {
  const source = Array.isArray(workbooks) ? workbooks.filter(Boolean) : [];
  const uniqueByKey = new Map();

  source.forEach((workbook) => {
    const key = toSafeString(workbook?.key || getWorkbookKey(workbook));
    if (!key || uniqueByKey.has(key)) {
      return;
    }

    uniqueByKey.set(key, {
      name: toSafeString(workbook?.name) || 'Workbook',
      path: toSafeString(workbook?.path),
      key
    });
  });

  const rows = Array.from(uniqueByKey.values());
  rows.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

  // Pin active workbook near the top
  const normalizedActiveWorkbookKey = toSafeString(activeWorkbookKey);
  if (normalizedActiveWorkbookKey) {
    const activeIndex = rows.findIndex((row) => row.key === normalizedActiveWorkbookKey);
    if (activeIndex > 0) {
      const [activeWorkbook] = rows.splice(activeIndex, 1);
      rows.unshift(activeWorkbook);
    }
  }

  // Pin PERSONAL.XLSB to the very top (always first)
  const personalIndex = rows.findIndex((row) => row.name.toUpperCase() === 'PERSONAL.XLSB');
  if (personalIndex > 0) {
    const [personal] = rows.splice(personalIndex, 1);
    rows.unshift(personal);
  }

  return rows;
}

export function resolveSelectedWorkbookKey({ requestedKey, workbooks, activeWorkbookKey }) {
  const source = Array.isArray(workbooks) ? workbooks : [];
  const normalizedRequested = toSafeString(requestedKey);
  const normalizedActive = toSafeString(activeWorkbookKey);

  if (normalizedRequested && source.some((workbook) => workbook.key === normalizedRequested)) {
    return normalizedRequested;
  }
  if (normalizedActive) {
    const activeWorkbook = source.find((workbook) => workbook.key === normalizedActive);
    if (activeWorkbook && toSafeString(activeWorkbook.name).toUpperCase() !== 'PERSONAL.XLSB') {
      return normalizedActive;
    }
  }
  // Skip PERSONAL.XLSB as default — it shows in the global macros section instead
  const fallback = source.find(
    (workbook) => toSafeString(workbook.name).toUpperCase() !== 'PERSONAL.XLSB'
  );
  return fallback?.key || source[0]?.key || '';
}

export function qualifyWorkbookNameForRun(workbookName) {
  const safeWorkbookName = toSafeString(workbookName);
  if (!safeWorkbookName) {
    return '';
  }
  if (/\s/.test(safeWorkbookName) || safeWorkbookName.includes("'")) {
    return `'${safeWorkbookName.replace(/'/g, "''")}'`;
  }
  return safeWorkbookName;
}

export function qualifyMacroFullName(workbookName, runTarget) {
  const safeRunTarget = toSafeString(runTarget);
  if (!safeRunTarget) {
    return '';
  }
  if (safeRunTarget.includes('!')) {
    return safeRunTarget;
  }

  const qualifiedWorkbook = qualifyWorkbookNameForRun(workbookName);
  if (!qualifiedWorkbook) {
    return safeRunTarget;
  }
  return `${qualifiedWorkbook}!${safeRunTarget}`;
}

export function namespaceMacrosForWorkbook(macros, workbook) {
  const source = Array.isArray(macros) ? macros : [];
  const workbookKey = getWorkbookKey(workbook) || 'workbook';

  return source.map((macro) => {
    const rawMacroId = String(macro?.id || `${macro?.module || 'module'}.${macro?.name || 'macro'}`);
    const runTarget = toSafeString(macro?.runTarget || `${macro?.module || ''}.${macro?.name || ''}`);
    return {
      ...macro,
      id: `${workbookKey}::macro::${rawMacroId}`,
      fullName: qualifyMacroFullName(workbook?.name, runTarget)
    };
  });
}

export function sortAllFilesModules(modules, activeWorkbookKey = '') {
  const source = Array.isArray(modules) ? modules.filter(Boolean) : [];
  const normalizedActiveWorkbookKey = toSafeString(activeWorkbookKey);
  const activeRows = [];
  const nonActiveRows = [];

  source.forEach((moduleItem) => {
    const workbookKey = getWorkbookKey({
      name: moduleItem?.workbookName,
      path: moduleItem?.workbookPath
    });
    if (normalizedActiveWorkbookKey && workbookKey === normalizedActiveWorkbookKey) {
      activeRows.push(moduleItem);
      return;
    }
    nonActiveRows.push(moduleItem);
  });

  const compareByName = (a, b) => {
    const aName = String(a?.name || '');
    const bName = String(b?.name || '');
    return aName.localeCompare(bName, undefined, { sensitivity: 'base' });
  };

  activeRows.sort(compareByName);
  nonActiveRows.sort((a, b) => {
    const workbookCompare = String(a?.workbookName || '').localeCompare(
      String(b?.workbookName || ''),
      undefined,
      { sensitivity: 'base' }
    );
    if (workbookCompare !== 0) {
      return workbookCompare;
    }
    return compareByName(a, b);
  });

  return [...activeRows, ...nonActiveRows];
}
