function toQuery(value) {
  return String(value || '').toLowerCase().trim();
}

export function getWorkbookContext(workbook) {
  const workbookName = workbook?.name ? String(workbook.name) : '';
  const workbookPath = workbook?.path ? String(workbook.path) : '';

  return {
    label: workbookName ? `Active workbook: ${workbookName}` : 'Active workbook unavailable',
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
    const savedShortcut = shortcutByMacroId[macro.id] || '';
    const haystack = `${macro?.name || ''} ${macro?.module || ''} ${savedShortcut}`.toLowerCase();
    return haystack.includes(normalizedQuery);
  });
}
