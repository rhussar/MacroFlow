const STANDARD_MODULE_TYPE_ID = 1;
const VBA_MODULE_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_]*$/;

export function normalizeModuleName(value) {
  return String(value || '').trim();
}

export function isStandardModule(moduleItem) {
  return Number(moduleItem?.typeId) === STANDARD_MODULE_TYPE_ID;
}

export function canShowModuleContextActions(moduleItem) {
  return Boolean(moduleItem) && isStandardModule(moduleItem) && Boolean(normalizeModuleName(moduleItem?.name));
}

export function isValidVbaModuleName(value) {
  return VBA_MODULE_NAME_PATTERN.test(normalizeModuleName(value));
}

export function shouldCommitModuleRename({ currentName, nextName }) {
  const normalizedCurrent = normalizeModuleName(currentName);
  const normalizedNext = normalizeModuleName(nextName);

  if (!normalizedCurrent || !normalizedNext) {
    return false;
  }

  return normalizedCurrent.toLowerCase() !== normalizedNext.toLowerCase();
}

export function buildWorkbookModuleRequest(moduleItem, fallbackWorkbook = null) {
  const moduleName = normalizeModuleName(moduleItem?.name);
  const workbookName = normalizeModuleName(moduleItem?.workbookName || fallbackWorkbook?.name);
  const workbookPath = normalizeModuleName(moduleItem?.workbookPath || fallbackWorkbook?.path);

  return {
    workbookName,
    workbookPath,
    moduleName
  };
}
