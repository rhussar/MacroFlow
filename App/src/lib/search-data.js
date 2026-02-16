const INTERNAL_MODULE_NAME = 'MacroFlow_Runtime';

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

export function normalizeWorkbook(workbookInfo, fallbackWorkbook = null) {
  const source = workbookInfo?.success
    ? { name: workbookInfo.name, path: workbookInfo.path }
    : fallbackWorkbook;

  if (!source) {
    return null;
  }

  const name = toSafeString(source.name);
  const path = toSafeString(source.path);
  if (!name && !path) {
    return null;
  }

  return {
    name: name || 'Active Workbook',
    path
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
      return moduleName && !isInternalModuleName(moduleName);
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
        runTarget
      };
    });
}

export function mapSearchError(message) {
  const rawMessage = toSafeString(message);
  const normalizedMessage = rawMessage.toUpperCase();

  if (normalizedMessage.includes('NO_EXCEL')) {
    return {
      status: 'no_excel',
      code: 'NO_EXCEL',
      message: 'Excel is not running. Open Excel, then select Refresh.'
    };
  }

  if (normalizedMessage.includes('NO_WORKBOOK')) {
    return {
      status: 'no_workbook',
      code: 'NO_WORKBOOK',
      message: 'Excel is open but no workbook is active. Open or create a workbook, then select Refresh.'
    };
  }

  return {
    status: 'error',
    code: normalizedMessage.includes('VBA_BLOCKED') ? 'VBA_BLOCKED' : 'UNKNOWN',
    message: rawMessage || 'Unable to load live workbook data.'
  };
}
