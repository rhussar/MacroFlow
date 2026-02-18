function toQuery(value) {
  return String(value || '').toLowerCase().trim();
}

function toWorkbookName(value) {
  return String(value || '').trim().toUpperCase();
}

export function filterModulesByQuery(modules, query) {
  const source = Array.isArray(modules) ? modules : [];
  const normalizedQuery = toQuery(query);

  return source.filter((moduleItem) => {
    const haystack = `${moduleItem?.name || ''} ${moduleItem?.type || ''} ${moduleItem?.workbookName || ''}`
      .toLowerCase();
    return haystack.includes(normalizedQuery);
  });
}

export function filterMacrosByQuery(macros, query, shortcutByMacroId = {}) {
  const source = Array.isArray(macros) ? macros : [];
  const normalizedQuery = toQuery(query);

  return source.filter((macro) => {
    const savedShortcut = String(shortcutByMacroId[macro.id] || '');
    const isUppercaseShortcutLetter = /^[A-Z]$/.test(savedShortcut);
    const shortcutTokens = savedShortcut
      ? `ctrl ${isUppercaseShortcutLetter ? 'shift ' : ''}${savedShortcut.toLowerCase()}`
      : '';
    const haystack = `${macro?.name || ''} ${macro?.module || ''} ${savedShortcut} ${shortcutTokens}`
      .toLowerCase();
    return haystack.includes(normalizedQuery);
  });
}

export function buildMacroRowUiModel(macro, source = 'active') {
  const fallbackId = `${macro?.module || 'module'}.${macro?.name || 'macro'}`;
  const macroId = String(macro?.id || fallbackId);
  return {
    uiId: `${source}::${macroId}`,
    source,
    macro
  };
}

export function selectActiveWorkbookMacros(macros, query, shortcutByMacroId = {}) {
  const filtered = filterMacrosByQuery(macros, query, shortcutByMacroId);
  return filtered.map((macro) => buildMacroRowUiModel(macro, 'active'));
}

export function selectPersonalGlobalMacros(macros, query) {
  const filtered = filterMacrosByQuery(macros, query, {});
  return filtered.map((macro) => buildMacroRowUiModel(macro, 'personal'));
}

export function selectPersonalGlobalSectionModel({
  workbookName,
  activeWorkbookName,
  rows,
  status = 'idle',
  workbookFound = false,
  error = null
}) {
  const normalizedActiveWorkbook = toWorkbookName(workbookName || activeWorkbookName);
  const normalizedPersonalWorkbook = 'PERSONAL.XLSB';
  const hidden = normalizedActiveWorkbook === normalizedPersonalWorkbook;
  const safeRows = Array.isArray(rows) ? rows : [];
  const count = safeRows.length;
  const isEmpty = count === 0;

  if (hidden) {
    return {
      hidden: true,
      count: 0,
      isEmpty: false,
      emptyMessage: ''
    };
  }

  let emptyMessage = '';
  if (isEmpty) {
    if (status === 'loading') {
      emptyMessage = 'Loading PERSONAL.XLSB macros...';
    } else if (status === 'error') {
      emptyMessage = error?.message
        ? String(error.message)
        : 'Could not load PERSONAL.XLSB macros.';
    } else if (!workbookFound) {
      emptyMessage = 'PERSONAL.XLSB is not open.';
    } else {
      emptyMessage = 'No global macros found in PERSONAL.XLSB.';
    }
  }

  return {
    hidden: false,
    count,
    isEmpty,
    emptyMessage
  };
}

