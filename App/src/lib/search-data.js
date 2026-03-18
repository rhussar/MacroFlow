const INTERNAL_MODULE_NAME = 'MacroFlow_Runtime';
const VBA_DOCUMENT_TYPE_ID = 100;

function toSafeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function toNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function isInternalModuleName(name) {
  return toSafeString(name).toLowerCase() === INTERNAL_MODULE_NAME.toLowerCase();
}

function isVbaDocumentObject(module) {
  const typeId = Number(module?.typeId);
  if (Number.isFinite(typeId) && typeId === VBA_DOCUMENT_TYPE_ID) {
    return true;
  }

  const typeName = toSafeString(module?.type).toLowerCase();
  return typeName === 'document';
}

export function normalizeWorkbook(workbookInfo, fallbackWorkbook = null) {
  const source = workbookInfo?.success
    ? { name: workbookInfo.name, path: workbookInfo.path, activeSheet: workbookInfo.activeSheet }
    : fallbackWorkbook;

  if (!source) {
    return null;
  }

  const name = toSafeString(source.name);
  const path = toSafeString(source.path);
  const activeSheet = toSafeString(source.activeSheet);
  if (!name && !path) {
    return null;
  }

  return {
    name: name || 'Active Workbook',
    path,
    activeSheet
  };
}

export function normalizeModules(apiModules = [], workbook = null) {
  const source = Array.isArray(apiModules) ? apiModules : [];
  const workbookName = toSafeString(workbook?.name);
  const workbookPath = toSafeString(workbook?.path);
  const workbookKey = workbookPath || workbookName || 'active-workbook';

  return source
    .filter((module) => {
      const moduleName = toSafeString(module?.name);
      return moduleName && !isInternalModuleName(moduleName) && !isVbaDocumentObject(module);
    })
    .map((module) => {
      const name = toSafeString(module.name);
      const type = toSafeString(module.type) || 'Unknown';
      const lineCount = toNumber(module.lineCount, 0);

      return {
        id: `${workbookKey}::module::${name}`,
        name,
        type,
        typeId: Number.isFinite(Number(module.typeId)) ? Number(module.typeId) : null,
        lineCount,
        workbookName,
        workbookPath
      };
    });
}

function isRunnableSubProcedure(procedure) {
  const kind = toSafeString(procedure?.kind);
  const scope = toSafeString(procedure?.scope);
  if (!/^sub\b/i.test(kind)) {
    return false;
  }

  if (!scope) {
    return true;
  }

  const normalizedScope = scope.toLowerCase();
  return normalizedScope === 'public' || normalizedScope === 'implicit';
}

export function normalizeMacros(apiProcedures = []) {
  const source = Array.isArray(apiProcedures) ? apiProcedures : [];

  return source
    .filter((procedure) => {
      const moduleName = toSafeString(procedure?.module);
      const procedureName = toSafeString(procedure?.name);
      if (!moduleName || !procedureName) {
        return false;
      }
      if (isInternalModuleName(moduleName)) {
        return false;
      }
      return isRunnableSubProcedure(procedure);
    })
    .map((procedure) => {
      const module = toSafeString(procedure.module);
      const name = toSafeString(procedure.name);
      const kind = toSafeString(procedure.kind) || 'Sub';
      const scope = toSafeString(procedure.scope) || 'Implicit';
      const runTarget = `${module}.${name}`;

      return {
        id: `${module}::${name}::${kind}::${scope}`,
        name,
        module,
        kind,
        scope,
        runTarget,
        fullName: runTarget
      };
    });
}

export function mapSearchError(message) {
  const rawMessage = toSafeString(message);
  const normalizedMessage = rawMessage.toUpperCase();

  if (normalizedMessage.includes('NO_VISIBLE_WINDOWS')) {
    return {
      status: 'excel_background',
      code: 'NO_VISIBLE_WINDOWS',
      message: 'Excel is running without a visible workbook window. MacroFlow will retry automatically.'
    };
  }

  if (normalizedMessage.includes('NO_EXCEL')) {
    return {
      status: 'no_excel',
      code: 'NO_EXCEL',
      message: 'Excel is not running. Open Excel and MacroFlow will retry automatically.'
    };
  }

  if (normalizedMessage.includes('MULTI_INSTANCE')) {
    return {
      status: 'multi_instance',
      code: 'MULTI_INSTANCE',
      message: 'Multiple Excel processes detected. Click on your workbook, then return here.'
    };
  }

  if (normalizedMessage.includes('NO_WORKBOOK')) {
    return {
      status: 'no_workbook',
      code: 'NO_WORKBOOK',
      message: 'Excel is open but no workbook is active. Open or create a workbook and MacroFlow will retry automatically.'
    };
  }

  if (normalizedMessage.includes('VBA_BLOCKED')) {
    return {
      status: 'error',
      code: 'VBA_BLOCKED',
      message: 'VBA project access is disabled. Enable it in File > Options > Trust Center > Macro Settings.'
    };
  }

  if (normalizedMessage.includes('COUNT') && normalizedMessage.includes('UNDEFINED')
    || normalizedMessage.includes('VBA PROJECT NOT ACCESSIBLE')) {
    return {
      status: 'error',
      code: 'VBA_LOCKED',
      message: 'This workbook is locked.'
    };
  }

  return {
    status: 'error',
    code: 'UNKNOWN',
    message: rawMessage || 'Unable to load live workbook data.'
  };
}
