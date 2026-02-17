import { filterModulesByQuery, filterMacrosByQuery } from './search-selectors';

/**
 * Build the unified Explorer item list from live search data.
 * Modules appear first, then macros.
 */
export function buildExplorerItems(modules, macros) {
  const safeModules = Array.isArray(modules) ? modules : [];
  const safeMacros = Array.isArray(macros) ? macros : [];

  const items = [];

  safeModules.forEach((mod) => {
    items.push({ ...mod, itemType: 'module' });
  });

  safeMacros.forEach((macro) => {
    items.push({ ...macro, itemType: 'macro' });
  });

  return items;
}

/**
 * Filter the unified Explorer item list by query using the shared selectors.
 */
export function filterExplorerItems(modules, macros, query, shortcutByMacroId = {}) {
  const filteredModules = filterModulesByQuery(modules, query);
  const filteredMacros = filterMacrosByQuery(macros, query, shortcutByMacroId);
  return buildExplorerItems(filteredModules, filteredMacros);
}

/**
 * Build metadata rows for the Explorer detail pane from a selected item.
 */
export function buildItemMetadata(item) {
  if (!item) return { rows: [], itemType: null };

  if (item.itemType === 'module') {
    return {
      itemType: 'module',
      rows: [
        { label: 'Name', value: item.name },
        { label: 'Type', value: item.type || 'Unknown' },
        { label: 'Lines', value: item.lineCount != null ? String(item.lineCount) : '-' },
        { label: 'Workbook', value: item.workbookName || '-' },
      ],
    };
  }

  return {
    itemType: 'macro',
    rows: [
      { label: 'Name', value: item.name },
      { label: 'Type', value: item.kind || 'Sub' },
      { label: 'Module', value: item.module || '-' },
      { label: 'Scope', value: item.scope || '-' },
    ],
  };
}
