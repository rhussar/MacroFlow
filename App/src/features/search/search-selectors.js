function toQuery(value) {
  return String(value || '').toLowerCase().trim();
}

export function getWorkbookContext(workbook) {
  const workbookName = workbook?.name ? String(workbook.name) : '';
  const workbookPath = workbook?.path ? String(workbook.path) : '';
  const activeSheetName = workbook?.activeSheet ? String(workbook.activeSheet) : '';
  const displayWorkbook = workbookName || 'Active Workbook';
  const displaySheet = activeSheetName || 'No active sheet';

  return {
    label: `🟩 ${displayWorkbook}  •  ${displaySheet}`,
    path: workbookPath
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
