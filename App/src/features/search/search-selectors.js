function toQuery(value) {
  return String(value || '').toLowerCase().trim();
}

const SEARCH_STATUS_COPY = {
  idle: {
    title: 'Loading workbook data',
    message: 'Connecting to the active Excel workbook.'
  },
  loading: {
    title: 'Loading workbook data',
    message: 'Refreshing modules and macros from Excel.'
  },
  no_excel: {
    title: 'Excel is not running',
    message: 'Open Excel. MacroFlow will retry automatically.'
  },
  no_workbook: {
    title: 'No active workbook',
    message: 'Open or create a workbook. MacroFlow will retry automatically.'
  },
  excel_background: {
    title: 'Excel background process detected',
    message: 'Excel is running without a visible workbook window. MacroFlow will retry automatically.'
  },
  multi_instance: {
    title: 'Multiple Excel instances detected',
    message: 'Click on your Excel workbook, then come back. MacroFlow will reconnect automatically.'
  },
  error: {
    title: 'Could not load workbook data',
    message: 'Something went wrong while reading workbook data. Retrying automatically.'
  }
};

function toWorkbookName(value) {
  return String(value || '').trim().toUpperCase();
}

export function getSearchStatusView(searchData) {
  const status = String(searchData?.status || 'idle');
  const config = SEARCH_STATUS_COPY[status] || SEARCH_STATUS_COPY.error;
  const message = status === 'error' && searchData?.error?.message
    ? String(searchData.error.message)
    : config.message;

  return {
    status,
    title: config.title,
    message,
    isLoading: status === 'idle' || status === 'loading'
  };
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

export function selectAllFilesModules(modules, query) {
  return filterModulesByQuery(modules, query);
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

export function selectPersonalGlobalMacros(macros, query, shortcutByMacroId = {}) {
  const filtered = filterMacrosByQuery(macros, query, shortcutByMacroId);
  return filtered.map((macro) => buildMacroRowUiModel(macro, 'personal'));
}

export function selectPersonalGlobalSectionModel({
  workbookName,
  activeWorkbookName,
  rows,
  totalMacros = 0,
  status = 'idle',
  workbookFound = false,
  fileExists = false,
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
  let action = null;
  if (isEmpty) {
    if (status === 'loading') {
      emptyMessage = 'Loading PERSONAL.XLSB macros...';
    } else if (status === 'error') {
      emptyMessage = error?.message
        ? String(error.message)
        : 'Could not load PERSONAL.XLSB macros.';
    } else if (Number(totalMacros) > 0) {
      emptyMessage = 'No global macros match this search.';
    } else if (!workbookFound) {
      if (!fileExists) {
        emptyMessage = 'PERSONAL.xlsb not found';
        action = 'create_file';
      } else {
        emptyMessage = 'PERSONAL.xlsb not open';
        action = 'open_file';
      }
    } else {
      emptyMessage = 'No macros found';
      action = 'create_global_macro';
    }
  }

  return {
    hidden: false,
    count,
    isEmpty,
    emptyMessage,
    action
  };
}
