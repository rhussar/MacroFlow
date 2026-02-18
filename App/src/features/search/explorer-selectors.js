import { PERSONAL_WORKBOOK_NAME } from './usePersonalMacros';

/**
 * @typedef {Object} TreeNode
 * @property {string}     id
 * @property {'workbook'|'module'|'macro'} nodeType
 * @property {string}     label
 * @property {Object|null} data     - Original domain object, null for inferred nodes
 * @property {TreeNode[]} children
 */

/**
 * Build the hierarchical Explorer tree from live search data.
 * Returns 0–2 workbook root nodes (active + optional PERSONAL).
 */
export function buildExplorerTree(searchData, personalState) {
  if (!searchData || searchData.status !== 'ready' || !searchData.workbook) {
    return [];
  }

  const workbook = searchData.workbook;
  const modules = Array.isArray(searchData.modules) ? searchData.modules : [];
  const macros = Array.isArray(searchData.macros) ? searchData.macros : [];

  // Group macros by module name
  const macrosByModule = new Map();
  for (const macro of macros) {
    const key = macro.module || '';
    if (!macrosByModule.has(key)) macrosByModule.set(key, []);
    macrosByModule.get(key).push(macro);
  }

  // Track which module names are covered by real module objects
  const knownModuleNames = new Set(modules.map((m) => m.name));

  // Build module nodes from real modules
  const moduleNodes = modules.map((mod) => ({
    id: mod.id,
    nodeType: 'module',
    label: mod.name,
    data: mod,
    children: (macrosByModule.get(mod.name) || []).map((macro) => ({
      id: macro.id,
      nodeType: 'macro',
      label: macro.name,
      data: macro,
      children: [],
    })),
  }));

  // Build synthetic module nodes for orphan macros
  for (const [moduleName, orphanMacros] of macrosByModule) {
    if (knownModuleNames.has(moduleName)) continue;
    const wbKey = String(workbook.path || workbook.name).trim();
    moduleNodes.push({
      id: `${wbKey}::module::${moduleName}`,
      nodeType: 'module',
      label: moduleName || 'Unknown Module',
      data: null,
      children: orphanMacros.map((macro) => ({
        id: macro.id,
        nodeType: 'macro',
        label: macro.name,
        data: macro,
        children: [],
      })),
    });
  }

  // Sort modules alphabetically
  moduleNodes.sort((a, b) => a.label.localeCompare(b.label));

  const activeWbNode = {
    id: `wb::${workbook.path || workbook.name}`,
    nodeType: 'workbook',
    label: workbook.name,
    data: workbook,
    children: moduleNodes,
  };

  const tree = [activeWbNode];

  // Build PERSONAL workbook node if applicable
  const activeNameUpper = String(workbook.name || '').trim().toUpperCase();
  const isPersonalActive = activeNameUpper === PERSONAL_WORKBOOK_NAME;

  if (
    !isPersonalActive &&
    personalState &&
    personalState.status === 'ready' &&
    Array.isArray(personalState.macros) &&
    personalState.macros.length > 0
  ) {
    const personalMacrosByModule = new Map();
    for (const macro of personalState.macros) {
      const key = macro.module || '';
      if (!personalMacrosByModule.has(key)) personalMacrosByModule.set(key, []);
      personalMacrosByModule.get(key).push(macro);
    }

    const personalModuleNodes = [];
    for (const [moduleName, pMacros] of personalMacrosByModule) {
      personalModuleNodes.push({
        id: `personal::module::${moduleName}`,
        nodeType: 'module',
        label: moduleName || 'Unknown Module',
        data: null,
        children: pMacros.map((macro) => ({
          id: `personal::${macro.id}`,
          nodeType: 'macro',
          label: macro.name,
          data: macro,
          children: [],
        })),
      });
    }

    personalModuleNodes.sort((a, b) => a.label.localeCompare(b.label));

    tree.push({
      id: 'wb::PERSONAL.XLSB',
      nodeType: 'workbook',
      label: 'PERSONAL.XLSB (Global Macros)',
      data: null,
      children: personalModuleNodes,
    });
  }

  return tree;
}

/**
 * Filter the Explorer tree by search query with bubble-up semantics.
 * Returns the filtered tree and a Set of all node IDs that should be expanded.
 */
export function filterExplorerTree(tree, query, shortcutByMacroId = {}) {
  const q = String(query || '').trim().toLowerCase();
  if (!q) {
    return { filteredTree: tree, matchedIds: new Set() };
  }

  const matchedIds = new Set();

  function filterNode(node) {
    if (node.nodeType === 'macro') {
      const shortcut = String(shortcutByMacroId[node.data?.id] || '');
      const isUpper = /^[A-Z]$/.test(shortcut);
      const shortcutTokens = shortcut
        ? `ctrl ${isUpper ? 'shift ' : ''}${shortcut.toLowerCase()}`
        : '';
      const haystack = `${node.label} ${node.data?.module || ''} ${shortcut} ${shortcutTokens}`.toLowerCase();
      if (haystack.includes(q)) {
        matchedIds.add(node.id);
        return node;
      }
      return null;
    }

    // Branch node (workbook or module)
    const selfMatch = node.label.toLowerCase().includes(q);
    const filteredChildren = node.children.map(filterNode).filter(Boolean);

    if (filteredChildren.length > 0 || selfMatch) {
      matchedIds.add(node.id);
      return {
        ...node,
        // Self-match shows all children; otherwise only show filtered
        children: selfMatch ? node.children : filteredChildren,
      };
    }

    return null;
  }

  const filteredTree = tree.map(filterNode).filter(Boolean);
  return { filteredTree, matchedIds };
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
          { label: 'Active Sheet', value: node.data?.activeSheet || '-' },
        ],
      };

    case 'module':
      return {
        nodeType: 'module',
        rows: [
          { label: 'Name', value: node.data?.name || node.label },
          { label: 'Type', value: node.data?.type || 'Unknown' },
          { label: 'Lines', value: node.data?.lineCount != null ? String(node.data.lineCount) : '-' },
          { label: 'Workbook', value: node.data?.workbookName || '-' },
        ],
      };

    case 'macro':
      return {
        nodeType: 'macro',
        rows: [
          { label: 'Name', value: node.data?.name || node.label },
          { label: 'Type', value: node.data?.kind || 'Sub' },
          { label: 'Module', value: node.data?.module || '-' },
          { label: 'Scope', value: node.data?.scope || '-' },
        ],
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
 * Return default expanded IDs: active workbook expanded, PERSONAL collapsed.
 */
export function getDefaultExpandedIds(tree) {
  const ids = new Set();
  if (tree.length > 0 && tree[0].nodeType === 'workbook') {
    ids.add(tree[0].id);
  }
  return ids;
}
