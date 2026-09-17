import { PERSONAL_WORKBOOK_NAME } from './usePersonalMacros.js';

/**
 * @typedef {Object} TreeNode
 * @property {string}     id
 * @property {'workbook'|'module'|'macro'} nodeType
 * @property {string}     label
 * @property {Object|null} data
 * @property {TreeNode[]} children
 */

function toSafeString(value) {
  return String(value || '').trim();
}

function toWorkbookKey(workbook) {
  return toSafeString(workbook?.path) || toSafeString(workbook?.name);
}

function buildModuleNodes({ workbook, modules = [], macros = [] }) {
  const macrosByModule = new Map();
  (Array.isArray(macros) ? macros : []).forEach((macro) => {
    const key = toSafeString(macro?.module);
    if (!macrosByModule.has(key)) {
      macrosByModule.set(key, []);
    }
    macrosByModule.get(key).push(macro);
  });

  const knownModuleNames = new Set(
    (Array.isArray(modules) ? modules : []).map((moduleItem) => toSafeString(moduleItem?.name))
  );

  const moduleNodes = (Array.isArray(modules) ? modules : []).map((moduleItem) => {
    const moduleName = toSafeString(moduleItem?.name);
    return {
      id: moduleItem.id,
      nodeType: 'module',
      label: moduleName,
      data: moduleItem,
      children: (macrosByModule.get(moduleName) || []).map((macro) => ({
        id: macro.id,
        nodeType: 'macro',
        label: macro.name,
        data: macro,
        children: []
      }))
    };
  });

  for (const [moduleName, orphanMacros] of macrosByModule) {
    if (knownModuleNames.has(moduleName)) {
      continue;
    }

    const workbookKey = toWorkbookKey(workbook) || 'workbook';
    moduleNodes.push({
      id: `${workbookKey}::module::${moduleName}`,
      nodeType: 'module',
      label: moduleName || 'Unknown Module',
      data: {
        name: moduleName || 'Unknown Module',
        workbookName: workbook?.name || '',
        workbookPath: workbook?.path || ''
      },
      children: orphanMacros.map((macro) => ({
        id: macro.id,
        nodeType: 'macro',
        label: macro.name,
        data: macro,
        children: []
      }))
    });
  }

  moduleNodes.sort((a, b) => a.label.localeCompare(b.label));
  return moduleNodes;
}

/**
 * Build the hierarchical Explorer tree from active workbook data plus open workbook/module context.
 * Supports both the legacy signature `(searchData, personalState)` and the object form.
 */
export function buildExplorerTree(searchDataOrOptions, personalStateArg) {
  const options = searchDataOrOptions && typeof searchDataOrOptions === 'object' && 'searchData' in searchDataOrOptions
    ? searchDataOrOptions
    : {
        searchData: searchDataOrOptions,
        personalState: personalStateArg,
        workbooks: [],
        allFilesModules: []
      };

  const searchData = options?.searchData;
  const personalState = options?.personalState;
  const openWorkbooks = Array.isArray(options?.workbooks) ? options.workbooks : [];
  const allFilesModules = Array.isArray(options?.allFilesModules) ? options.allFilesModules : [];
  const showModules = options?.showModules !== false;

  if (!searchData || searchData.status !== 'ready' || !searchData.workbook) {
    return [];
  }

  const activeWorkbook = searchData.workbook;
  const activeWorkbookKey = toWorkbookKey(activeWorkbook);
  const activeModules = Array.isArray(searchData.modules) ? searchData.modules : [];
  const activeMacros = Array.isArray(searchData.macros) ? searchData.macros : [];

  const workbookMap = new Map();
  const orderedWorkbookKeys = [];
  let personalRegistered = false;
  const registerWorkbook = (workbook) => {
    // Prevent duplicate PERSONAL.XLSB entries when the placeholder was
    // registered with name-only key but the real workbook has a path key.
    if (personalRegistered && toSafeString(workbook?.name).toUpperCase() === PERSONAL_WORKBOOK_NAME) {
      // Update the existing placeholder with the real path if we now have it
      const realPath = toSafeString(workbook?.path);
      if (realPath) {
        for (const [key, wb] of workbookMap) {
          if (toSafeString(wb.name).toUpperCase() === PERSONAL_WORKBOOK_NAME && !wb.path) {
            wb.path = realPath;
          }
        }
      }
      return;
    }

    const workbookKey = toWorkbookKey(workbook);
    if (!workbookKey || workbookMap.has(workbookKey)) {
      return;
    }

    workbookMap.set(workbookKey, {
      ...workbook,
      name: toSafeString(workbook?.name) || 'Workbook',
      path: toSafeString(workbook?.path)
    });
    orderedWorkbookKeys.push(workbookKey);
  };

  // Register PERSONAL.XLSB first so it always appears at the top.
  // Show a placeholder immediately (even while loading) so the tree
  // doesn't shift when personal data arrives later.
  const personalStatus = toSafeString(personalState?.status);
  const personalConfirmedMissing =
    personalStatus === 'ready' && !personalState?.workbookFound &&
    !(Array.isArray(personalState?.macros) && personalState.macros.length > 0);
  if (personalState && !personalConfirmedMissing) {
    registerWorkbook({
      name: PERSONAL_WORKBOOK_NAME,
      path: toSafeString(personalState?.workbook?.path || personalState?.workbookPath)
    });
    personalRegistered = true;
  }

  registerWorkbook(activeWorkbook);
  openWorkbooks.forEach((workbook) => registerWorkbook(workbook));

  const modulesByWorkbook = new Map();
  const registerModule = (moduleItem) => {
    const workbookKey = toWorkbookKey({
      name: moduleItem?.workbookName,
      path: moduleItem?.workbookPath
    });
    if (!workbookKey) {
      return;
    }

    if (!modulesByWorkbook.has(workbookKey)) {
      modulesByWorkbook.set(workbookKey, new Map());
    }

    const moduleId = toSafeString(moduleItem?.id);
    if (!moduleId) {
      return;
    }
    modulesByWorkbook.get(workbookKey).set(moduleId, moduleItem);
  };

  allFilesModules.forEach((moduleItem) => registerModule(moduleItem));
  activeModules.forEach((moduleItem) => registerModule(moduleItem));

  const tree = orderedWorkbookKeys.map((workbookKey) => {
    const workbook = workbookMap.get(workbookKey);
    const workbookNameUpper = toSafeString(workbook?.name).toUpperCase();
    const workbookModules = Array.from(modulesByWorkbook.get(workbookKey)?.values() || []);
    const workbookMacros = workbookKey === activeWorkbookKey
      ? activeMacros
      : (workbookNameUpper === PERSONAL_WORKBOOK_NAME && personalState?.status === 'ready'
        ? (Array.isArray(personalState?.macros) ? personalState.macros : [])
        : []);

    const children = showModules
      ? buildModuleNodes({ workbook, modules: workbookModules, macros: workbookMacros })
      : (Array.isArray(workbookMacros) ? workbookMacros : []).map((macro) => ({
          id: macro.id,
          nodeType: 'macro',
          label: macro.name,
          data: macro,
          children: []
        }));

    return {
      id: `wb::${workbookKey}`,
      nodeType: 'workbook',
      label: workbookNameUpper === PERSONAL_WORKBOOK_NAME && workbookKey !== activeWorkbookKey
        ? 'PERSONAL.XLSB (Global Macros)'
        : workbook.name,
      data: workbook,
      children
    };
  });

  return tree;
}

/**
 * Build metadata rows for the detail pane from a selected tree node.
 */
export function buildNodeMetadata(node) {
  if (!node) return { rows: [], nodeType: null };

  switch (node.nodeType) {
    case 'workbook':
      return {
        nodeType: 'workbook',
        rows: [
          { label: 'Name', value: node.data?.name || node.label },
          { label: 'Path', value: node.data?.path || '-' },
          { label: 'Active Sheet', value: node.data?.activeSheet || '-' }
        ]
      };

    case 'module':
      return {
        nodeType: 'module',
        rows: [
          { label: 'Name', value: node.data?.name || node.label },
          { label: 'Type', value: node.data?.type || 'Unknown' },
          { label: 'Lines', value: node.data?.lineCount != null ? String(node.data.lineCount) : '-' },
          { label: 'Workbook', value: node.data?.workbookName || '-' }
        ]
      };

    case 'macro':
      return {
        nodeType: 'macro',
        rows: [
          { label: 'Name', value: node.data?.name || node.label },
          { label: 'Type', value: node.data?.kind || 'Sub' },
          { label: 'Module', value: node.data?.module || '-' },
          { label: 'Scope', value: node.data?.scope || '-' }
        ]
      };

    default:
      return { rows: [], nodeType: null };
  }
}

/**
 * Count all module + macro nodes in the tree (excludes workbook roots).
 */
export function countTreeItems(tree) {
  let count = 0;
  function walk(nodes) {
    for (const node of nodes) {
      if (node.nodeType === 'module' || node.nodeType === 'macro') count++;
      if (node.children.length > 0) walk(node.children);
    }
  }
  walk(tree);
  return count;
}

/**
 * Return default expanded IDs for the Explorer tree.
 * Workbook roots start collapsed until the user expands them.
 */
export function getDefaultExpandedIds() {
  return new Set();
}
