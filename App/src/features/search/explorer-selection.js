function toSafeString(value) {
  return String(value || '').trim();
}

export function resolveInitialModuleNode(tree, { moduleId = '', moduleName = '' } = {}) {
  const source = Array.isArray(tree) ? tree : [];
  const normalizedId = toSafeString(moduleId);
  const normalizedName = toSafeString(moduleName).toLowerCase();

  const walk = (nodes) => {
    for (const node of nodes) {
      if (node?.nodeType === 'module') {
        if (normalizedId && String(node.id || '').trim() === normalizedId) {
          return node;
        }
        const nodeName = toSafeString(node?.data?.name || node?.label).toLowerCase();
        if (!normalizedId && normalizedName && nodeName === normalizedName) {
          return node;
        }
      }
      if (Array.isArray(node?.children) && node.children.length > 0) {
        const childMatch = walk(node.children);
        if (childMatch) {
          return childMatch;
        }
      }
    }
    return null;
  };

  return walk(source);
}
