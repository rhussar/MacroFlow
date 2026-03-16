/**
 * ExcelBridge - COM bridge for Excel automation.
 *
 * This is the only file that should touch Excel via COM.
 * All UI/IPC layers call into this class.
 */

const { execSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const winax = require('winax');
const logger = require('./logger');

function parseTasklistRows(output) {
  return String(output || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line && !line.toUpperCase().startsWith('INFO:'));
}

function parseCsvQuotedColumns(line) {
  const columns = [];
  const regex = /"((?:[^"]|"")*)"/g;
  let match;
  while ((match = regex.exec(String(line || '')))) {
    columns.push(String(match[1] || '').replace(/""/g, '"'));
  }
  return columns;
}

// VBA Component Types (vbext_ComponentType)
const VBA_COMPONENT_TYPE = {
  STANDARD_MODULE: 1,
  CLASS_MODULE: 2,
  FORM: 3,
  DOCUMENT: 100
};

const VBA_COMPONENT_NAME = {
  [VBA_COMPONENT_TYPE.STANDARD_MODULE]: 'Standard Module',
  [VBA_COMPONENT_TYPE.CLASS_MODULE]: 'Class Module',
  [VBA_COMPONENT_TYPE.FORM]: 'UserForm',
  [VBA_COMPONENT_TYPE.DOCUMENT]: 'Document'
};

const PERSONAL_WORKBOOK_NAME = 'PERSONAL.XLSB';
const XLSB_FILE_FORMAT = 50;
const VBA_MODULE_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_]*$/;
const MACROFLOW_ADDIN_WORKBOOK_NAME = 'MacroFlow.xlam';
const WINDOWLESS_GRACE_MS = 2500;
const WINDOWLESS_QUIT_COOLDOWN_MS = 10000;
const WINDOWLESS_MAX_QUIT_ATTEMPTS = 1;

class ExcelBridge {
  constructor() {
    this._proceduresCache = null;
    this._focusHelper = null;
    this._multiInstanceCacheTimestamp = 0;
    this._multiInstanceCacheResult = null;
    this._isShuttingDown = false;
    this._attachAttemptSeq = 0;
    this._windowlessFirstDetectedAt = 0;
    this._windowlessLastQuitAttemptAt = 0;
    this._windowlessQuitAttempts = 0;
    this._windowlessProcessSignature = '';
  }

  setFocusHelper(client) {
    this._focusHelper = client || null;
  }

  setShuttingDown(value = true) {
    this._isShuttingDown = Boolean(value);
    logger.info('[ExcelBridge] shutdown latch updated', { shuttingDown: this._isShuttingDown });
    if (this._isShuttingDown) {
      this.clearComCache();
    }
  }

  clearComCache() {
    this._proceduresCache = null;
    this._multiInstanceCacheTimestamp = 0;
    this._multiInstanceCacheResult = null;
  }

  isShuttingDown() {
    return this._isShuttingDown;
  }

  _resetWindowlessLifecycle(reason = '') {
    const hadWindowlessState =
      this._windowlessFirstDetectedAt > 0 ||
      this._windowlessLastQuitAttemptAt > 0 ||
      this._windowlessQuitAttempts > 0;

    if (hadWindowlessState) {
      logger.info('[ExcelBridge] windowless_cleared', {
        reason: String(reason || '').trim() || 'reset',
        firstDetectedAt: this._windowlessFirstDetectedAt,
        lastQuitAttemptAt: this._windowlessLastQuitAttemptAt,
        quitAttempts: this._windowlessQuitAttempts,
        processSignature: this._windowlessProcessSignature
      });
    }

    this._windowlessFirstDetectedAt = 0;
    this._windowlessLastQuitAttemptAt = 0;
    this._windowlessQuitAttempts = 0;
    this._windowlessProcessSignature = '';
  }

  _safeRelease(...objects) {
    const releasable = objects.filter(Boolean);
    if (releasable.length < 1) {
      return;
    }
    try {
      winax.release(...releasable);
    } catch {
      // Ignore release failures.
    }
  }

  _listExcelProcessIds() {
    try {
      const output = execSync('tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH', {
        windowsHide: true,
        timeout: 3000,
        encoding: 'utf8'
      });
      const rows = parseTasklistRows(output);
      return rows
        .map((line) => {
          const match = line.match(/^"EXCEL\.EXE","(\d+)"/i);
          return match ? Number(match[1]) : NaN;
        })
        .filter((value) => Number.isFinite(value));
    } catch {
      return [];
    }
  }

  getExcelProcessIds() {
    return this._listExcelProcessIds();
  }

  _hasVisibleExcelWindow(processIds = []) {
    const processIdSet = new Set(
      (Array.isArray(processIds) ? processIds : [])
        .map((value) => Number(value))
        .filter((value) => Number.isFinite(value))
    );

    try {
      const output = execSync('tasklist /FI "IMAGENAME eq EXCEL.EXE" /V /FO CSV /NH', {
        windowsHide: true,
        timeout: 3000,
        encoding: 'utf8'
      });
      const rows = parseTasklistRows(output);
      for (const line of rows) {
        const columns = parseCsvQuotedColumns(line);
        if (columns.length < 2) {
          continue;
        }

        const imageName = String(columns[0] || '').trim().toUpperCase();
        if (imageName !== 'EXCEL.EXE') {
          continue;
        }

        const pid = Number(columns[1]);
        if (processIdSet.size > 0) {
          if (!Number.isFinite(pid) || !processIdSet.has(pid)) {
            continue;
          }
        }

        const windowTitle = String(columns[8] || columns[columns.length - 1] || '').trim();
        if (windowTitle && windowTitle.toUpperCase() !== 'N/A') {
          return true;
        }
      }

      return false;
    } catch {
      // Fail open if visibility check is unavailable so we don't block real sessions.
      return true;
    }
  }

  _buildProcessSignature(processIds = []) {
    if (!Array.isArray(processIds) || processIds.length < 1) {
      return '';
    }

    return processIds
      .map((value) => Number(value))
      .filter((value) => Number.isFinite(value))
      .sort((a, b) => a - b)
      .join(',');
  }

  // ===========================================================================
  // CONNECTION
  // ===========================================================================

  /**
   * Connect to a running Excel instance (never creates a new one).
   * Returns a fresh COM reference for each call.
   * @returns {object} Excel.Application COM object
   * @throws {Error} If Excel is not running
   */
  getApp(options = {}) {
    const { activate = true } = options;
    const attempt = ++this._attachAttemptSeq;

    logger.debug('[ExcelBridge] COM attach requested', {
      attempt,
      activate,
      shuttingDown: this._isShuttingDown
    });

    if (this._isShuttingDown) {
      logger.warn('[ExcelBridge] COM attach blocked by shutdown latch', { attempt, activate });
      throw new Error('APP_SHUTTING_DOWN: MacroFlow is closing and Excel operations are paused.');
    }

    const processIds = this._listExcelProcessIds();
    if (processIds.length === 0) {
      this._resetWindowlessLifecycle('no_processes_preflight');
      logger.debug('[ExcelBridge] COM attach preflight found no Excel processes', { attempt });
      throw new Error('NO_EXCEL: Excel is not running. Please open Excel first.');
    }
    const processSignature = this._buildProcessSignature(processIds);
    const hasVisibleWindowPreflight = this._hasVisibleExcelWindow(processIds);
    if (!hasVisibleWindowPreflight) {
      logger.info('[ExcelBridge] preflight_no_visible_windows', {
        attempt,
        processCount: processIds.length,
        processSignature
      });
      throw new Error(
        'NO_VISIBLE_WINDOWS: Excel process found but has no visible workbook windows. ' +
        'It may be shutting down. Please reopen Excel.'
      );
    }

    let excel = null;
    try {
      excel = new winax.Object('Excel.Application', { activate });
      if (!excel) {
        throw new Error('Excel application not found');
      }
      const afterAttachProcessIds = this._listExcelProcessIds();
      const attachProcessSetChanged =
        processIds.some((pid) => !afterAttachProcessIds.includes(pid)) ||
        afterAttachProcessIds.some((pid) => !processIds.includes(pid));

      // Verify Excel has at least one open window. A windowless EXCEL.EXE is
      // a ghost process that is shutting down — attaching to it would keep it
      // alive and lock the .xlam add-in file.
      let windowsProxy = null;
      try {
        windowsProxy = excel.Windows;
        const windowCount = windowsProxy ? Number(windowsProxy.Count) : 0;
        if (windowCount < 1) {
          if (attachProcessSetChanged) {
            logger.warn('[ExcelBridge] attach_process_cycle_changed_while_windowless', {
              attempt,
              processIdsBeforeAttach: processIds,
              processIdsAfterAttach: afterAttachProcessIds
            });
            this._resetWindowlessLifecycle('process_cycle_changed_while_windowless');
            this._safeRelease(windowsProxy, excel);
            windowsProxy = null;
            excel = null;
            if (typeof global.gc === 'function') {
              try { global.gc(); } catch { /* best-effort */ }
            }
            throw new Error('NO_EXCEL: Excel is shutting down. Please wait for it to close and reopen Excel.');
          }

          const nowMs = Date.now();
          if (
            this._windowlessProcessSignature &&
            this._windowlessProcessSignature !== processSignature
          ) {
            logger.info('[ExcelBridge] windowless_cycle_changed', {
              attempt,
              previousProcessSignature: this._windowlessProcessSignature,
              processSignature
            });
            this._windowlessFirstDetectedAt = 0;
            this._windowlessLastQuitAttemptAt = 0;
            this._windowlessQuitAttempts = 0;
          }

          if (this._windowlessFirstDetectedAt < 1) {
            this._windowlessProcessSignature = processSignature;
            this._windowlessFirstDetectedAt = nowMs;
            logger.info('[ExcelBridge] windowless_detected', {
              attempt,
              processCount: processIds.length,
              processSignature
            });
          }

          const elapsedSinceFirstMs = nowMs - this._windowlessFirstDetectedAt;
          const elapsedSinceLastQuitMs = this._windowlessLastQuitAttemptAt > 0
            ? nowMs - this._windowlessLastQuitAttemptAt
            : Number.POSITIVE_INFINITY;
          const shouldAttemptQuit =
            elapsedSinceFirstMs >= WINDOWLESS_GRACE_MS &&
            this._windowlessQuitAttempts < WINDOWLESS_MAX_QUIT_ATTEMPTS &&
            elapsedSinceLastQuitMs >= WINDOWLESS_QUIT_COOLDOWN_MS;

          // Best-effort graceful cleanup only after grace period.
          if (shouldAttemptQuit) {
            const quitAttemptNumber = this._windowlessQuitAttempts + 1;
            logger.info('[ExcelBridge] windowless_quit_attempted', {
              attempt,
              quitAttemptNumber,
              processCount: processIds.length,
              elapsedSinceFirstMs
            });
            this._windowlessLastQuitAttemptAt = nowMs;
            this._windowlessQuitAttempts = quitAttemptNumber;
            try {
              excel.DisplayAlerts = false;
              excel.Quit();
            } catch {
              // Best-effort; the instance may already be partially torn down.
            }
          }

          this._safeRelease(windowsProxy, excel);
          windowsProxy = null;
          excel = null;
          // Force immediate GC to release any transient COM proxies created
          // during the attach + quit sequence.
          if (typeof global.gc === 'function') {
            try { global.gc(); } catch { /* best-effort */ }
          }
          throw new Error(
            'NO_VISIBLE_WINDOWS: Excel process found but has no visible workbook windows. ' +
            'It may be shutting down. Please reopen Excel.'
          );
        }
      } finally {
        this._safeRelease(windowsProxy);
      }

      logger.debug('[ExcelBridge] COM attach succeeded', {
        attempt,
        processCount: processIds.length,
        processCountAfterAttach: afterAttachProcessIds.length
      });
      if (afterAttachProcessIds.length > processIds.length) {
        logger.warn('[ExcelBridge] Excel process count increased after COM attach', {
          attempt,
          processIdsBeforeAttach: processIds,
          processIdsAfterAttach: afterAttachProcessIds
        });
      }
      this._resetWindowlessLifecycle('attach_succeeded');
      return excel;
    } catch (error) {
      // Ensure the COM reference is released if any validation step threw.
      this._safeRelease(excel);
      excel = null;

      const remainingProcessIds = this._listExcelProcessIds();
      logger.warn('[ExcelBridge] COM attach failed', {
        attempt,
        processCountBefore: processIds.length,
        processCountAfter: remainingProcessIds.length,
        error: String(error?.message || error || 'Unknown COM attach error')
      });
      if (remainingProcessIds.length === 0) {
        this._resetWindowlessLifecycle('no_processes_after_failure');
        throw new Error('NO_EXCEL: Excel is not running. Please open Excel first.');
      }
      const errorMessage = String(error?.message || '');
      if (errorMessage.includes('NO_EXCEL') || errorMessage.includes('NO_VISIBLE_WINDOWS')) {
        throw error;
      }
      if (remainingProcessIds.length > 1) {
        throw new Error(
          'MULTI_INSTANCE: Multiple Excel processes detected. ' +
          'Click on your Excel workbook, then return to MacroFlow.'
        );
      }
      throw new Error(
        `NO_WORKBOOK: Unable to connect to the active workbook in Excel. ${error.message || ''}`.trim()
      );
    }
  }

  _withExcelApp(operation, options = {}) {
    let excel = null;
    try {
      excel = this.getApp(options);
      return operation(excel);
    } finally {
      this._safeRelease(excel);
    }
  }

  _getActiveWorkbookFromApp(excel) {
    const workbook = excel?.ActiveWorkbook;
    if (!workbook) {
      if (this._hasMultipleExcelProcesses()) {
        throw new Error(
          'MULTI_INSTANCE: Multiple Excel processes detected. ' +
          'Click on your Excel workbook, then return to MacroFlow.'
        );
      }
      throw new Error('NO_WORKBOOK: No workbook is open. Please open or create a workbook.');
    }
    return workbook;
  }

  _withActiveWorkbook(operation, options = {}) {
    return this._withExcelApp((excel) => {
      const workbook = this._getActiveWorkbookFromApp(excel);
      try {
        return operation({ excel, workbook });
      } finally {
        this._safeRelease(workbook);
      }
    }, options);
  }

  /**
   * Get the active workbook.
   * @returns {object} Workbook COM object
   * @throws {Error} If no workbook is open
   */
  getActiveWorkbook(options = {}) {
    const excel = options.excel || this.getApp(options);
    return this._getActiveWorkbookFromApp(excel);
  }

  _hasMultipleExcelProcesses() {
    const now = Date.now();
    if (this._multiInstanceCacheResult !== null && now - this._multiInstanceCacheTimestamp < 2000) {
      return this._multiInstanceCacheResult;
    }

    const result = this._listExcelProcessIds().length > 1;
    this._multiInstanceCacheTimestamp = now;
    this._multiInstanceCacheResult = result;
    return result;
  }

  hasMultipleExcelProcesses() {
    return this._hasMultipleExcelProcesses();
  }

  /**
   * Get the VBA project from the active workbook.
   * @returns {object} VBProject COM object
   * @throws {Error} If VBA access is not trusted
   */
  getVBProject(options = {}) {
    const workbook = options.workbook || this.getActiveWorkbook(options);
    return this._getVBProjectForWorkbook(workbook);
  }

  _getVBProjectForWorkbook(workbook) {
    try {
      const vbProject = workbook.VBProject;
      if (!vbProject) {
        throw new Error('VBA project not accessible');
      }
      return vbProject;
    } catch (error) {
      throw new Error(
        'VBA_BLOCKED: Trust access to VBA project is disabled. ' +
          'Go to File > Options > Trust Center > Trust Center Settings > ' +
          'Macro Settings > Check "Trust access to the VBA project object model".'
      );
    }
  }

  // ===========================================================================
  // HELPERS (PURE)
  // ===========================================================================

  _getWorkbookCacheKey(workbook) {
    const path = workbook && workbook.FullName ? String(workbook.FullName) : '';
    const name = workbook && workbook.Name ? String(workbook.Name) : '';
    return `${path}::${name}`;
  }

  _describeWorkbook(workbook) {
    return {
      name: workbook.Name,
      path: workbook.FullName
    };
  }

  _getPersonalWorkbookPath() {
    const appData = String(process.env.APPDATA || '').trim();
    if (!appData) {
      throw new Error('APPDATA environment variable is not available.');
    }

    return path.join(appData, 'Microsoft', 'Excel', 'XLSTART', PERSONAL_WORKBOOK_NAME);
  }

  _ensureDirectoryExists(filePath) {
    const directory = path.dirname(String(filePath || '').trim());
    if (!directory) {
      throw new Error('Unable to resolve PERSONAL.XLSB directory.');
    }
    fs.mkdirSync(directory, { recursive: true });
  }

  _describeWorkbookWithActiveSheet(workbook) {
    let activeSheet = '';
    let activeSheetRef = null;
    let workbookSheets = null;
    const sheetRefs = [];
    const sheets = [];

    try {
      try {
        activeSheetRef = workbook.ActiveSheet;
        activeSheet = activeSheetRef ? String(activeSheetRef.Name) : '';
      } catch {
        activeSheet = '';
      }

      workbookSheets = workbook?.Sheets;
      const count = workbookSheets ? Number(workbookSheets.Count) : 0;
      for (let i = 1; i <= count; i += 1) {
        const sheet = workbookSheets.Item(i);
        sheetRefs.push(sheet);
        sheets.push(String(sheet.Name));
      }
    } finally {
      this._safeRelease(activeSheetRef, ...sheetRefs, workbookSheets);
    }

    return {
      ...this._describeWorkbook(workbook),
      activeSheet,
      sheets
    };
  }

  _findOpenWorkbookByName(excel, workbookName) {
    return this._findOpenWorkbook(excel, { workbookName });
  }

  _findOpenWorkbook(excel, { workbookName, workbookPath } = {}) {
    const normalizedName = String(workbookName || '').trim().toLowerCase();
    const normalizedPath = String(workbookPath || '').trim().toLowerCase();
    if (!normalizedName && !normalizedPath) {
      return null;
    }

    const workbooks = excel?.Workbooks;
    const count = workbooks ? Number(workbooks.Count) : 0;
    if (!Number.isFinite(count) || count < 1) {
      return null;
    }

    const nonMatchingRefs = [];
    let matchedWorkbook = null;

    try {
      if (normalizedPath) {
        for (let i = 1; i <= count; i++) {
          const workbook = workbooks.Item(i);
          if (!workbook) {
            continue;
          }
          const candidatePath = String(workbook.FullName || '').trim().toLowerCase();
          if (candidatePath === normalizedPath) {
            matchedWorkbook = workbook;
            break;
          }
          nonMatchingRefs.push(workbook);
        }
      }

      if (!matchedWorkbook && normalizedName) {
        for (let i = 1; i <= count; i++) {
          const workbook = workbooks.Item(i);
          if (!workbook) {
            continue;
          }
          const candidateName = String(workbook.Name || '').trim().toLowerCase();
          if (candidateName === normalizedName) {
            matchedWorkbook = workbook;
            break;
          }
          nonMatchingRefs.push(workbook);
        }
      }

      return matchedWorkbook;
    } finally {
      this._safeRelease(...nonMatchingRefs, workbooks);
    }
  }

  _findComponentByName(vbProject, name) {
    let components = null;
    let found = null;
    try {
      components = vbProject.VBComponents;
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        if (component.Name === name) {
          found = component;
          break;
        }
        this._safeRelease(component);
      }
      return found;
    } finally {
      this._safeRelease(components);
    }
  }

  _qualifyWorkbookName(name) {
    const safe = String(name || '');
    if (!safe) {
      return '';
    }
    if (/\s/.test(safe) || safe.includes("'")) {
      return `'${safe.replace(/'/g, "''")}'`;
    }
    return safe;
  }

  _normalizeMacroNameForWorkbook(macroName, workbook, procedures = []) {
    if (!macroName) {
      return '';
    }

    const trimmed = String(macroName).trim();
    if (!trimmed) {
      return '';
    }

    if (trimmed.includes('!')) {
      return trimmed;
    }

    const workbookName = this._qualifyWorkbookName(workbook?.Name);
    if (!workbookName) {
      return trimmed;
    }

    if (trimmed.includes('.')) {
      return `${workbookName}!${trimmed}`;
    }

    const sourceProcedures = Array.isArray(procedures) ? procedures : [];
    const matches = sourceProcedures.filter(
      (proc) => String(proc?.name || '').toLowerCase() === trimmed.toLowerCase()
    );
    if (matches.length === 1) {
      return `${workbookName}!${matches[0].module}.${matches[0].name}`;
    }

    return trimmed;
  }

  _normalizeMacroName(macroName, workbook) {
    const procedures = workbook
      ? this._listMacroProceduresForWorkbook(workbook)
      : this._listMacroProcedures().procedures;
    return this._normalizeMacroNameForWorkbook(macroName, workbook, procedures);
  }

  _componentTypeName(typeId) {
    return VBA_COMPONENT_NAME[typeId] || `Unknown (${typeId})`;
  }

  _isStandardModuleComponent(component) {
    return Number(component?.Type) === VBA_COMPONENT_TYPE.STANDARD_MODULE;
  }

  _isValidVbaModuleName(name) {
    return VBA_MODULE_NAME_PATTERN.test(String(name || '').trim());
  }

  _listModulesForWorkbook(workbook, vbProject) {
    const components = vbProject.VBComponents;
    const modules = [];
    const componentRefs = [];
    const codeModuleRefs = [];

    try {
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        componentRefs.push(component);
        const codeModule = component.CodeModule;
        if (codeModule) {
          codeModuleRefs.push(codeModule);
        }
        const lineCount = codeModule ? codeModule.CountOfLines : 0;
        modules.push({
          name: component.Name,
          typeId: component.Type,
          type: this._componentTypeName(component.Type),
          lineCount
        });
      }

      return modules;
    } finally {
      this._safeRelease(...codeModuleRefs, ...componentRefs, components);
    }
  }

  _listProceduresForWorkbook(workbook, vbProject) {
    const components = vbProject.VBComponents;
    const workbookKey = this._getWorkbookCacheKey(workbook);
    const previousModuleEntries =
      this._proceduresCache && this._proceduresCache.workbookKey === workbookKey
        ? this._proceduresCache.moduleEntries
        : {};
    const nextModuleEntries = {};
    const procedures = [];
    const componentRefs = [];
    const codeModuleRefs = [];

    try {
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        componentRefs.push(component);
        const moduleName = String(component.Name);
        const codeModule = component.CodeModule;
        if (!codeModule) {
          continue;
        }
        codeModuleRefs.push(codeModule);

        const lineCount = Number(codeModule.CountOfLines) || 0;
        if (lineCount < 1) {
          continue;
        }

        const cachedModule = previousModuleEntries[moduleName];
        if (cachedModule && cachedModule.lineCount === lineCount) {
          nextModuleEntries[moduleName] = cachedModule;
          cachedModule.procedures.forEach((proc) => {
            procedures.push({ ...proc });
          });
          continue;
        }

        const codeText = codeModule.Lines(1, lineCount);
        const parsed = this._parseProcedures(codeText);
        const moduleProcedures = parsed.map((proc) => ({
          module: moduleName,
          ...proc
        }));

        nextModuleEntries[moduleName] = {
          lineCount,
          procedures: moduleProcedures
        };

        moduleProcedures.forEach((proc) => {
          procedures.push({ ...proc });
        });
      }

      this._proceduresCache = {
        workbookKey,
        moduleEntries: nextModuleEntries
      };

      return procedures;
    } finally {
      this._safeRelease(...codeModuleRefs, ...componentRefs, components);
    }
  }

  _parseProcedures(codeText) {
    const procedures = [];
    const regex =
      /^\s*(Public|Private|Friend)?\s*(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)/gim;
    let match;
    while ((match = regex.exec(codeText)) !== null) {
      const scope = match[1] ? match[1] : 'Implicit';
      const isStatic = Boolean(match[2]);
      const kind = match[3].replace(/\s+/g, ' ');
      const name = match[5];
      procedures.push({
        name,
        kind,
        scope,
        isStatic
      });
    }
    return procedures;
  }

  _normalizeValue(value) {
    if (value === null || value === undefined) {
      return null;
    }
    if (value instanceof Date) {
      return value.toISOString();
    }
    if (Array.isArray(value)) {
      return value.map((item) => this._normalizeValue(item));
    }
    if (typeof value === 'object') {
      if (typeof value.valueOf === 'function') {
        const primitive = value.valueOf();
        if (primitive === null || primitive === undefined) {
          return null;
        }
        if (typeof primitive !== 'object') {
          return primitive;
        }
      }
      return String(value);
    }
    return value;
  }

  _columnLetter(index) {
    let dividend = index;
    let columnName = '';
    while (dividend > 0) {
      const modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
  }

  _classifyValue(value) {
    if (value === null || value === undefined || value === '') {
      return 'empty';
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      return 'number';
    }
    if (typeof value === 'boolean') {
      return 'boolean';
    }
    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (!trimmed) {
        return 'empty';
      }
      if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
        return 'number-text';
      }
      if (!Number.isNaN(Date.parse(trimmed))) {
        return 'date-text';
      }
      return 'text';
    }
    return 'other';
  }

  _normalizeRangeValues(values) {
    const normalizeRow = (row) => {
      if (!Array.isArray(row)) {
        return [this._normalizeValue(row)];
      }
      const trimmedRow = row.length > 0 && row[0] === undefined ? row.slice(1) : row;
      return trimmedRow.map((cell) => this._normalizeValue(cell));
    };

    if (Array.isArray(values)) {
      const hasRowArrays = Array.isArray(values[0]) || Array.isArray(values[1]);
      if (hasRowArrays) {
        const rows = values.length > 0 && values[0] === undefined ? values.slice(1) : values;
        return rows.map((row) => normalizeRow(row));
      }
      return [normalizeRow(values)];
    }

    return [[this._normalizeValue(values)]];
  }

  _normalizeVbaCode(code) {
    return String(code || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  }

  _hashFnv1aHex(text) {
    const bytes = Buffer.from(String(text || ''), 'utf8');
    let hash = 0x811c9dc5;

    for (let i = 0; i < bytes.length; i += 1) {
      hash ^= bytes[i];
      hash = Math.imul(hash, 0x01000193) >>> 0;
    }

    return hash.toString(16).padStart(8, '0');
  }

  _buildCodeSignature(code) {
    const normalizedCode = this._normalizeVbaCode(code);
    const lineCount = normalizedCode ? normalizedCode.split('\n').length : 0;
    return {
      code: normalizedCode,
      lineCount,
      hash: this._hashFnv1aHex(normalizedCode)
    };
  }

  _readCodeModuleText(codeModule) {
    const lineCount = Number(codeModule?.CountOfLines) || 0;
    if (lineCount < 1) {
      return '';
    }
    return String(codeModule.Lines(1, lineCount) || '');
  }

  _getModuleCodeForWorkbook(vbProject, moduleName) {
    const normalizedModuleName = String(moduleName || '').trim();
    if (!normalizedModuleName) {
      return {
        success: false,
        moduleFound: false,
        moduleName: '',
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        code: '',
        message: 'Module name is required.'
      };
    }

    let component = null;
    let codeModule = null;

    try {
      component = this._findComponentByName(vbProject, normalizedModuleName);
      if (!component) {
        const emptySignature = this._buildCodeSignature('');
        return {
          success: true,
          moduleFound: false,
          moduleName: normalizedModuleName,
          lineCount: emptySignature.lineCount,
          hash: emptySignature.hash,
          code: '',
          message: `Module "${normalizedModuleName}" was not found.`
        };
      }

      codeModule = component.CodeModule;
      const code = this._readCodeModuleText(codeModule);
      const signature = this._buildCodeSignature(code);
      return {
        success: true,
        moduleFound: true,
        moduleName: normalizedModuleName,
        lineCount: signature.lineCount,
        hash: signature.hash,
        code: signature.code
      };
    } finally {
      this._safeRelease(codeModule, component);
    }
  }

  _setModuleCodeForWorkbook(vbProject, moduleName, code, options = {}) {
    const { createIfMissing = false } = options;
    const normalizedModuleName = String(moduleName || '').trim();
    if (!normalizedModuleName) {
      return {
        success: false,
        moduleFound: false,
        moduleName: '',
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        message: 'Module name is required.'
      };
    }

    const signature = this._buildCodeSignature(code);
    let component = null;
    let codeModule = null;
    let componentsProxy = null;

    try {
      component = this._findComponentByName(vbProject, normalizedModuleName);
      const hasExisting = Boolean(component);

      if (!component && !createIfMissing) {
        return {
          success: false,
          moduleFound: false,
          moduleName: normalizedModuleName,
          lineCount: signature.lineCount,
          hash: signature.hash,
          message: `Module "${normalizedModuleName}" was not found.`
        };
      }

      if (!component) {
        componentsProxy = vbProject.VBComponents;
        component = componentsProxy.Add(VBA_COMPONENT_TYPE.STANDARD_MODULE);
        component.Name = normalizedModuleName;
      }

      codeModule = component.CodeModule;
      const existingLineCount = Number(codeModule?.CountOfLines) || 0;
      if (existingLineCount > 0) {
        codeModule.DeleteLines(1, existingLineCount);
      }
      if (signature.code) {
        codeModule.InsertLines(1, signature.code);
      }

      return {
        success: true,
        moduleFound: hasExisting,
        moduleName: normalizedModuleName,
        lineCount: signature.lineCount,
        hash: signature.hash
      };
    } finally {
      this._safeRelease(codeModule, component, componentsProxy);
    }
  }

  _getRuntimeModuleCode() {
    return [
      'Option Explicit',
      '',
      'Public Function MacroFlow_RunMacro(ByVal macroName As String) As String',
      '    On Error GoTo Handler',
      '    Application.Run macroName',
      '    MacroFlow_RunMacro = "OK"',
      '    Exit Function',
      'Handler:',
      '    MacroFlow_RunMacro = "ERR|" & Err.Number & "|" & Err.Description & "|" & Err.Source',
      'End Function'
    ].join('\n');
  }

  _ensureRuntimeModule(workbook) {
    const vbProject = this._getVBProjectForWorkbook(workbook);
    const moduleName = 'MacroFlow_Runtime';
    const code = this._getRuntimeModuleCode();
    let existing = null;
    let newModule = null;
    let codeModule = null;
    let componentsProxy = null;

    try {
      existing = this._findComponentByName(vbProject, moduleName);

      if (existing) {
        codeModule = existing.CodeModule;
        if (codeModule.CountOfLines > 0) {
          const existingCode = codeModule.Lines(1, codeModule.CountOfLines);
          if (existingCode.includes('MacroFlow_RunMacro')) {
            return;
          }
          codeModule.DeleteLines(1, codeModule.CountOfLines);
        }
        codeModule.AddFromString(code);
        return;
      }

      componentsProxy = vbProject.VBComponents;
      newModule = componentsProxy.Add(VBA_COMPONENT_TYPE.STANDARD_MODULE);
      newModule.Name = moduleName;
      codeModule = newModule.CodeModule;
      codeModule.AddFromString(code);
    } finally {
      this._safeRelease(codeModule, existing, newModule, componentsProxy, vbProject);
    }
  }

  _buildAddinRuntimeTarget() {
    const addinName = this._qualifyWorkbookName(MACROFLOW_ADDIN_WORKBOOK_NAME);
    return `${addinName}!MacroFlow_Runtime.MacroFlow_RunMacro`;
  }

  _buildWorkbookRuntimeTarget(workbook) {
    const workbookName = this._qualifyWorkbookName(workbook?.Name);
    return `${workbookName}!MacroFlow_Runtime.MacroFlow_RunMacro`;
  }

  _parseRuntimeTrapResult(result) {
    if (typeof result !== 'string' || !result.startsWith('ERR|')) {
      return null;
    }

    const [, number, description, source] = result.split('|');
    return {
      success: false,
      message: `VBA error ${number}: ${description}`,
      error: {
        number: Number(number),
        description: description || '',
        source: source || ''
      }
    };
  }

  _isRuntimeUnavailableError(error) {
    const message = String(error?.message || error || '').toLowerCase();

    // Typical Excel COM text when the macro entry point cannot be found.
    if (message.includes('cannot run the macro')) {
      return true;
    }
    if (message.includes('the macro may not be available in this workbook')) {
      return true;
    }

    // Typical invoke/dispatch failure surfaces for missing method/member.
    if (message.includes('0x80020006')) {
      return true;
    }
    if (message.includes('0x80020003')) {
      return true;
    }

    // VBA runtime 1004 style surface for missing macro target.
    if (message.includes('0x800a03ec')) {
      return true;
    }

    return false;
  }

  _withWorkbookAtPath(path, callback) {
    return this._withExcelApp((excel) => {
      const state = {
        screenUpdating: excel.ScreenUpdating,
        displayAlerts: excel.DisplayAlerts,
        enableEvents: excel.EnableEvents
      };
      let workbook = null;
      let previousWorkbook = null;
      let workbookWindow = null;
      let workbooksProxy = null;
      let windowsProxy = null;

      try {
        previousWorkbook = excel.ActiveWorkbook;
        excel.ScreenUpdating = false;
        excel.DisplayAlerts = false;
        excel.EnableEvents = false;

        workbooksProxy = excel.Workbooks;
        workbook = workbooksProxy.Open(path, 0, true);
        try {
          windowsProxy = workbook.Windows;
          workbookWindow = windowsProxy.Item(1);
          workbookWindow.Visible = false;
        } catch (error) {
          // Ignore if window visibility cannot be changed.
        }

        return callback(workbook, excel);
      } finally {
        if (workbook) {
          try {
            workbook.Close(false);
          } catch (error) {
            // Ignore close errors.
          }
        }
        try {
          excel.ScreenUpdating = state.screenUpdating;
          excel.DisplayAlerts = state.displayAlerts;
          excel.EnableEvents = state.enableEvents;
        } catch (error) {
          // Ignore restore errors.
        }
        if (previousWorkbook) {
          try {
            previousWorkbook.Activate();
          } catch (error) {
            // Ignore activation errors.
          }
        }
        this._safeRelease(workbookWindow, windowsProxy, workbook, previousWorkbook, workbooksProxy);
      }
    });
  }

  _prepareMacroRun(excel, workbook) {
    try {
      workbook.Activate();
    } catch (error) {
      // Ignore if activation fails.
    }

    try {
      const activeWindow = excel.ActiveWindow;
      try {
        if (activeWindow) {
          activeWindow.Activate();
        }
      } finally {
        this._safeRelease(activeWindow);
      }
    } catch (error) {
      // Ignore window activation issues.
    }

    let ready = true;
    try {
      ready = Boolean(excel.Ready);
    } catch (error) {
      ready = true;
    }

    return {
      ready,
      message: ready ? '' : 'Excel is busy or in edit mode. Finish editing a cell to run macros.'
    };
  }

  _buildSelectionContext(excel) {
    let selectionAddress = '';
    let activeCellAddress = '';
    let activeCellValue = null;
    let selectionInTable = false;
    let tableInfo = null;

    let selection = null;
    let selectionListObject = null;
    let selectionListRange = null;

    try {
      selection = excel.Selection;
      if (selection) {
        selectionAddress = String(selection.Address).replace(/\$/g, '');
        try {
          selectionListObject = selection.ListObject;
          if (selectionListObject) {
            selectionListRange = selectionListObject.Range;
            selectionInTable = true;
            tableInfo = {
              name: String(selectionListObject.Name),
              range: String(selectionListRange.Address).replace(/\$/g, '')
            };
          }
        } catch (error) {
          // Not in a table.
        }
      }
    } catch (error) {
      // Ignore selection errors.
    } finally {
      this._safeRelease(selectionListRange, selectionListObject, selection);
    }

    let activeCell = null;
    let activeCellListObject = null;
    let activeCellListRange = null;

    try {
      activeCell = excel.ActiveCell;
      if (activeCell) {
        activeCellAddress = String(activeCell.Address).replace(/\$/g, '');
        activeCellValue = this._normalizeValue(activeCell.Value2);
        if (!selectionInTable) {
          try {
            activeCellListObject = activeCell.ListObject;
            if (activeCellListObject) {
              activeCellListRange = activeCellListObject.Range;
              selectionInTable = true;
              tableInfo = {
                name: String(activeCellListObject.Name),
                range: String(activeCellListRange.Address).replace(/\$/g, '')
              };
            }
          } catch (error) {
            // Not in a table.
          }
        }
      }
    } catch (error) {
      // Ignore active cell errors.
    } finally {
      this._safeRelease(activeCellListRange, activeCellListObject, activeCell);
    }

    return {
      address: selectionAddress,
      activeCell: {
        address: activeCellAddress,
        value: activeCellValue
      },
      inTable: selectionInTable,
      table: tableInfo
    };
  }

  _collectWorksheetMetadata({ excel, workbook, sheet, includeSelection }) {
    const buildRangeFromCells = (startRow, startColumn, endRow, endColumn) => {
      let startCell = null;
      let endCell = null;
      try {
        startCell = sheet.Cells(startRow, startColumn);
        endCell = sheet.Cells(endRow, endColumn);
        return sheet.Range(startCell, endCell);
      } finally {
        this._safeRelease(startCell, endCell);
      }
    };

    let usedRange = null;
    let usedRows = null;
    let usedColumns = null;
    let previewRange = null;
    let headerRange = null;
    let sampleRange = null;
    let workbookSheets = null;
    const sheetRefs = [];
    let activeSheetRef = null;

    try {
      usedRange = sheet.UsedRange;
      usedRows = usedRange ? usedRange.Rows : null;
      usedColumns = usedRange ? usedRange.Columns : null;

      const startRow = usedRange.Row;
      const startColumn = usedRange.Column;
      const totalRows = Number(usedRows.Count);
      const totalColumns = Number(usedColumns.Count);

      const SAMPLE_ROWS = 50;
      const SAMPLE_COLS = 25;

      const previewRows = Math.min(totalRows, SAMPLE_ROWS);
      const previewCols = Math.min(totalColumns, SAMPLE_COLS);

      previewRange = buildRangeFromCells(
        startRow,
        startColumn,
        startRow + previewRows - 1,
        startColumn + previewCols - 1
      );
      const previewValues = previewRange.Value2;
      const preview = this._normalizeRangeValues(previewValues);

      headerRange = buildRangeFromCells(
        startRow,
        startColumn,
        startRow,
        startColumn + previewCols - 1
      );
      const headerValues = this._normalizeRangeValues(headerRange.Value2);
      const headers = headerValues[0] || [];

      const headerAddressMap = {};
      for (let colIndex = 0; colIndex < previewCols; colIndex++) {
        const columnNumber = startColumn + colIndex;
        const address = `${this._columnLetter(columnNumber)}${startRow}`;
        headerAddressMap[address] = headers[colIndex] ?? null;
      }

      let formulaGrid = [];
      try {
        formulaGrid = this._normalizeRangeValues(previewRange.Formula);
      } catch (error) {
        formulaGrid = [];
      }

      const sampleRowCount = Math.max(0, Math.min(4, totalRows - 1));
      let sampleRows = [];
      const sampleRowStart = startRow + 1;
      if (sampleRowCount > 0) {
        sampleRange = buildRangeFromCells(
          sampleRowStart,
          startColumn,
          sampleRowStart + sampleRowCount - 1,
          startColumn + previewCols - 1
        );
        sampleRows = this._normalizeRangeValues(sampleRange.Value2);
      }

      const columns = [];
      for (let colIndex = 0; colIndex < previewCols; colIndex++) {
        const columnNumber = startColumn + colIndex;
        const columnLetter = this._columnLetter(columnNumber);
        const header = headers[colIndex] !== undefined ? headers[colIndex] : '';

        const typeCounts = {
          empty: 0,
          number: 0,
          boolean: 0,
          text: 0,
          'number-text': 0,
          'date-text': 0,
          other: 0
        };
        const examples = [];
        const exampleSet = new Set();
        let nonEmpty = 0;
        let numericMin = null;
        let numericMax = null;
        let formulaCells = 0;

        const analysisRows = sampleRows.length ? sampleRows : preview.slice(1);
        for (let rowIndex = 0; rowIndex < analysisRows.length; rowIndex++) {
          const value = analysisRows[rowIndex]?.[colIndex];
          const type = this._classifyValue(value);
          typeCounts[type] += 1;
          if (type !== 'empty') {
            nonEmpty += 1;
            const example = value === null || value === undefined ? '' : String(value);
            if (example && !exampleSet.has(example) && examples.length < 6) {
              examples.push(example);
              exampleSet.add(example);
            }
          }
          if (type === 'number') {
            if (numericMin === null || value < numericMin) numericMin = value;
            if (numericMax === null || value > numericMax) numericMax = value;
          }

          const formulaCell = formulaGrid[rowIndex + 1]?.[colIndex];
          if (typeof formulaCell === 'string' && formulaCell.startsWith('=')) {
            formulaCells += 1;
          }
        }

        const typeSummaryParts = Object.entries(typeCounts)
          .filter(([key, count]) => key !== 'empty' && count > 0)
          .sort((a, b) => b[1] - a[1])
          .map(([key, count]) => `${key} ${count}`);
        if (formulaCells > 0) {
          typeSummaryParts.push(`formulas ${formulaCells}`);
        }
        const typeSummary = typeSummaryParts.length ? typeSummaryParts.join(' | ') : 'empty';

        columns.push({
          index: columnNumber,
          column: columnLetter,
          header,
          nonEmpty,
          empty: typeCounts.empty,
          typeSummary,
          examples,
          numericMin,
          numericMax,
          formulaCells
        });
      }

      const sheetNames = [];
      let activeSheetName = '';
      workbookSheets = workbook.Sheets;
      const sheetCount = workbookSheets ? Number(workbookSheets.Count) : 0;
      for (let i = 1; i <= sheetCount; i++) {
        const sheetItem = workbookSheets.Item(i);
        sheetRefs.push(sheetItem);
        sheetNames.push(String(sheetItem.Name));
      }

      try {
        activeSheetRef = workbook.ActiveSheet;
        if (activeSheetRef) {
          activeSheetName = String(activeSheetRef.Name);
        }
      } catch (error) {
        activeSheetName = '';
      }

      const selectionContext = includeSelection ? this._buildSelectionContext(excel) : null;

      const structuralContext = {
        workbook: this._describeWorkbook(workbook),
        sheets: sheetNames,
        activeSheet: activeSheetName
      };

      const dataContext = {
        usedRange: {
          address: String(usedRange.Address),
          startRow,
          startColumn,
          rows: totalRows,
          columns: totalColumns
        },
        headers,
        headersByAddress: headerAddressMap,
        sampleRows: {
          startRow: sampleRowCount ? sampleRowStart : null,
          endRow: sampleRowCount ? sampleRowStart + sampleRowCount - 1 : null,
          rows: sampleRows
        },
        columns
      };

      const llmContext = {
        structural: {
          workbookName: structuralContext.workbook.name,
          workbookPath: structuralContext.workbook.path,
          worksheetNames: structuralContext.sheets,
          activeSheet: structuralContext.activeSheet
        },
        data: {
          usedRange: dataContext.usedRange,
          headersByAddress: dataContext.headersByAddress,
          sampleRows: dataContext.sampleRows,
          columns: dataContext.columns.map((column) => ({
            column: column.column,
            header: column.header,
            typeSummary: column.typeSummary,
            nonEmpty: column.nonEmpty,
            numericMin: column.numericMin,
            numericMax: column.numericMax,
            examples: column.examples,
            formulaCells: column.formulaCells
          }))
        },
        selection: selectionContext
      };

      return {
        sheet: {
          name: String(sheet.Name),
          index: Number(sheet.Index)
        },
        structuralContext,
        dataContext,
        selectionContext,
        llmContext,
        preview: {
          rows: preview,
          truncated: totalRows > previewRows || totalColumns > previewCols
        }
      };
    } finally {
      this._safeRelease(
        activeSheetRef,
        ...sheetRefs,
        workbookSheets,
        sampleRange,
        headerRange,
        previewRange,
        usedColumns,
        usedRows,
        usedRange
      );
    }
  }

  // ===========================================================================
  // WORKBOOK INFO
  // ===========================================================================

  /**
   * Get info about the active workbook.
   * @returns {{ success: boolean, name: string, path: string, activeSheet: string, sheets: string[] }}
   */
  getWorkbookInfo() {
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        let activeSheet = '';
        let activeSheetRef = null;
        let sheetsCollection = null;
        const sheetRefs = [];
        const sheets = [];

        try {
          try {
            activeSheetRef = workbook.ActiveSheet;
            activeSheet = activeSheetRef ? String(activeSheetRef.Name) : '';
          } catch (error) {
            activeSheet = '';
          }
          sheetsCollection = workbook.Sheets;
          const sheetCount = sheetsCollection ? Number(sheetsCollection.Count) : 0;
          for (let i = 1; i <= sheetCount; i++) {
            const sheet = sheetsCollection.Item(i);
            sheetRefs.push(sheet);
            sheets.push(sheet.Name);
          }
        } finally {
          this._safeRelease(activeSheetRef, ...sheetRefs, sheetsCollection);
        }

        return {
          success: true,
          name: workbook.Name,
          path: workbook.FullName,
          activeSheet,
          sheets
        };
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Get all open workbooks in Excel.
   * @returns {{ success: boolean, workbooks: Array<{ name: string, path: string }>, message?: string }}
   */
  getOpenWorkbooks() {
    try {
      return this._withExcelApp((excel) => {
        const workbooks = [];
        const workbookRefs = [];
        let workbookCollection = null;

        try {
          workbookCollection = excel.Workbooks;
          const count = workbookCollection ? Number(workbookCollection.Count) : 0;
          for (let i = 1; i <= count; i++) {
            const workbook = workbookCollection.Item(i);
            workbookRefs.push(workbook);
            workbooks.push(this._describeWorkbook(workbook));
          }
        } finally {
          this._safeRelease(...workbookRefs, workbookCollection);
        }

        return { success: true, workbooks };
      });
    } catch (error) {
      if (error.message.includes('NO_EXCEL')) {
        return { success: false, workbooks: [], message: 'Excel not found' };
      }
      return { success: false, workbooks: [], message: error.message };
    }
  }

  /**
   * Return PERSONAL.XLSB status based on XLSTART path and open workbook state.
   * @returns {{
   *   success: boolean,
   *   workbookFound: boolean,
   *   workbook: { name: string, path: string } | null,
   *   fileExists: boolean,
   *   workbookPath: string,
   *   message?: string
   * }}
   */
  getPersonalWorkbookStatus(options = {}) {
    const { activate = true } = options;
    let workbookPath = '';
    let fileExists = false;

    try {
      workbookPath = this._getPersonalWorkbookPath();
      fileExists = fs.existsSync(workbookPath);
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        message: error.message
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: PERSONAL_WORKBOOK_NAME,
          workbookPath
        });

        if (!workbook) {
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            fileExists,
            workbookPath,
            message: fileExists
              ? 'PERSONAL.XLSB is not open.'
              : 'PERSONAL.XLSB was not found in XLSTART.'
          };
        }

        try {
          return {
            success: true,
            workbookFound: true,
            workbook: this._describeWorkbook(workbook),
            fileExists: true,
            workbookPath
          };
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        fileExists,
        workbookPath,
        message: error.message
      };
    }
  }

  /**
   * Return PERSONAL.XLSB status, procedures, and shortcut audit in one COM session.
   * @returns {{
   *   success: boolean,
   *   workbookFound: boolean,
   *   workbook: { name: string, path: string } | null,
   *   fileExists: boolean,
   *   workbookPath: string,
   *   procedures: Array,
   *   shortcutAudit: { success: boolean, shortcuts: Array, unmapped: Array, note?: string, message?: string },
   *   message?: string
   * }}
   */
  getPersonalWorkbookContext(options = {}) {
    const { activate = true } = options;
    let workbookPath = '';
    let fileExists = false;

    try {
      workbookPath = this._getPersonalWorkbookPath();
      fileExists = fs.existsSync(workbookPath);
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        procedures: [],
        shortcutAudit: { success: false, shortcuts: [], unmapped: [], message: error.message },
        message: error.message
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: PERSONAL_WORKBOOK_NAME,
          workbookPath
        });

        if (!workbook) {
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            fileExists,
            workbookPath,
            procedures: [],
            shortcutAudit: { success: true, shortcuts: [], unmapped: [] },
            message: fileExists
              ? 'PERSONAL.XLSB is not open.'
              : 'PERSONAL.XLSB was not found in XLSTART.'
          };
        }

        let vbProject = null;
        try {
          vbProject = this._getVBProjectForWorkbook(workbook);
          const procedures = this._listProceduresForWorkbook(workbook, vbProject);
          let shortcutAudit = {
            success: true,
            shortcuts: [],
            unmapped: [],
            note: 'Excel does not expose global shortcut listings. Only MacroFlow-tracked shortcuts are available.'
          };
          try {
            shortcutAudit = this._auditShortcutsForWorkbook(workbook);
          } catch (error) {
            shortcutAudit = {
              success: false,
              shortcuts: [],
              unmapped: [],
              message: String(error?.message || 'Shortcut audit failed.')
            };
          }

          return {
            success: true,
            workbookFound: true,
            workbook: this._describeWorkbook(workbook),
            fileExists: true,
            workbookPath,
            procedures,
            shortcutAudit
          };
        } finally {
          this._safeRelease(vbProject, workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        fileExists,
        workbookPath,
        procedures: [],
        shortcutAudit: { success: false, shortcuts: [], unmapped: [], message: error.message },
        message: error.message
      };
    }
  }

  /**
   * Open PERSONAL.XLSB from XLSTART in the currently connected Excel instance.
   * @returns {{
   *   success: boolean,
   *   workbookFound: boolean,
   *   opened: boolean,
   *   alreadyOpen: boolean,
   *   workbook: { name: string, path: string } | null,
   *   fileExists: boolean,
   *   workbookPath: string,
   *   message?: string
   * }}
   */
  openPersonalWorkbook(options = {}) {
    const { activate = true } = options;
    let workbookPath = '';
    let fileExists = false;

    try {
      workbookPath = this._getPersonalWorkbookPath();
      fileExists = fs.existsSync(workbookPath);
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        opened: false,
        alreadyOpen: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        message: error.message
      };
    }

    if (!fileExists) {
      return {
        success: false,
        workbookFound: false,
        opened: false,
        alreadyOpen: false,
        workbook: null,
        fileExists: false,
        workbookPath,
        message: 'PERSONAL.XLSB was not found in XLSTART.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        let workbook = this._findOpenWorkbook(excel, {
          workbookName: PERSONAL_WORKBOOK_NAME,
          workbookPath
        });

        if (workbook) {
          try {
            workbook.Activate();
          } catch {
            // Ignore activation failures.
          }

          try {
            return {
              success: true,
              workbookFound: true,
              opened: false,
              alreadyOpen: true,
              workbook: this._describeWorkbook(workbook),
              fileExists: true,
              workbookPath,
              message: 'PERSONAL.XLSB is already open.'
            };
          } finally {
            this._safeRelease(workbook);
          }
        }

        let workbooksProxy = null;
        try {
          workbooksProxy = excel.Workbooks;
          workbook = workbooksProxy.Open(workbookPath);
        } catch (openError) {
          this._safeRelease(workbooksProxy);
          throw openError;
        }
        try {
          try {
            workbook.Activate();
          } catch {
            // Ignore activation failures.
          }

          return {
            success: true,
            workbookFound: true,
            opened: true,
            alreadyOpen: false,
            workbook: this._describeWorkbook(workbook),
            fileExists: true,
            workbookPath,
            message: 'Opened PERSONAL.XLSB.'
          };
        } finally {
          this._safeRelease(workbook, workbooksProxy);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        opened: false,
        alreadyOpen: false,
        workbook: null,
        fileExists,
        workbookPath,
        message: error.message
      };
    }
  }

  /**
   * Create PERSONAL.XLSB in XLSTART and open it in Excel.
   * @returns {{
   *   success: boolean,
   *   created: boolean,
   *   opened: boolean,
   *   workbookFound: boolean,
   *   workbook: { name: string, path: string } | null,
   *   fileExists: boolean,
   *   workbookPath: string,
   *   message?: string
   * }}
   */
  createPersonalWorkbook(options = {}) {
    const { activate = true } = options;
    let workbookPath = '';
    let fileExists = false;

    try {
      workbookPath = this._getPersonalWorkbookPath();
      this._ensureDirectoryExists(workbookPath);
      fileExists = fs.existsSync(workbookPath);
    } catch (error) {
      return {
        success: false,
        created: false,
        opened: false,
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        message: error.message
      };
    }

    if (fileExists) {
      const openResult = this.openPersonalWorkbook({ activate });
      return {
        ...openResult,
        created: false
      };
    }

    try {
      return this._withExcelApp((excel) => {
        let workbook = null;
        let workbooksProxy = null;
        const previousDisplayAlerts = excel.DisplayAlerts;

        try {
          excel.DisplayAlerts = false;
          workbooksProxy = excel.Workbooks;
          workbook = workbooksProxy.Add();
          workbook.SaveAs(workbookPath, XLSB_FILE_FORMAT);

          try {
            workbook.Activate();
          } catch {
            // Ignore activation failures.
          }

          return {
            success: true,
            created: true,
            opened: true,
            workbookFound: true,
            workbook: this._describeWorkbook(workbook),
            fileExists: true,
            workbookPath,
            message: 'Created and opened PERSONAL.XLSB.'
          };
        } finally {
          try {
            excel.DisplayAlerts = previousDisplayAlerts;
          } catch {
            // Ignore alert restore failures.
          }
          this._safeRelease(workbook, workbooksProxy);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        created: false,
        opened: false,
        workbookFound: false,
        workbook: null,
        fileExists,
        workbookPath,
        message: error.message
      };
    }
  }

  /**
   * List worksheets in the active workbook with basic UsedRange stats.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, sheets: Array }}
   */
  listWorksheets() {
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        let activeSheetRef = null;
        let activeSheetName = '';
        try {
          activeSheetRef = workbook.ActiveSheet;
          activeSheetName = activeSheetRef ? String(activeSheetRef.Name) : '';
        } catch (error) {
          activeSheetName = '';
        }
        const sheets = [];
        const sheetRefs = [];
        const rangeRefs = [];
        const rowRefs = [];
        const colRefs = [];
        let workbookSheets = null;

        try {
          workbookSheets = workbook.Sheets;
          const sheetCount = workbookSheets ? Number(workbookSheets.Count) : 0;
          for (let i = 1; i <= sheetCount; i++) {
            const sheet = workbookSheets.Item(i);
            sheetRefs.push(sheet);
            let usedRange = null;
            let usedRows = null;
            let usedCols = null;
            try {
              usedRange = sheet.UsedRange;
              if (usedRange) {
                rangeRefs.push(usedRange);
                usedRows = usedRange.Rows;
                usedCols = usedRange.Columns;
                rowRefs.push(usedRows);
                colRefs.push(usedCols);
              }
            } catch (error) {
              usedRange = null;
            }
            sheets.push({
              name: String(sheet.Name),
              index: i,
              visible: Number(sheet.Visible),
              active: activeSheetName ? activeSheetName === String(sheet.Name) : false,
              usedRange: usedRange
                ? {
                    address: String(usedRange.Address),
                    rows: Number(usedRows.Count),
                    columns: Number(usedCols.Count)
                  }
                : null
            });
          }
        } finally {
          this._safeRelease(activeSheetRef, ...colRefs, ...rowRefs, ...rangeRefs, ...sheetRefs, workbookSheets);
        }

        return {
          success: true,
          workbook: this._describeWorkbook(workbook),
          sheets
        };
      });
    } catch (error) {
      return { success: false, sheets: [], message: error.message };
    }
  }

  /**
   * Build workbook + sheet metadata for LLM prompting.
   * @param {{ sheetName?: string }} options
   * @returns {{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object, selectionContext?: object, llmContext?: object }}
   */
  getWorksheetMetadata(options = {}) {
    try {
      return this._withActiveWorkbook(({ excel, workbook }) => {
        let sheetsProxy = null;
        let sheet = null;
        if (options.sheetName) {
          sheetsProxy = workbook.Sheets;
          sheet = sheetsProxy.Item(options.sheetName);
        } else {
          sheet = workbook.ActiveSheet;
        }

        try {
          const metadata = this._collectWorksheetMetadata({
            excel,
            workbook,
            sheet,
            includeSelection: true
          });

          return {
            success: true,
            message: `Metadata read for "${sheet.Name}"`,
            ...metadata
          };
        } finally {
          this._safeRelease(sheetsProxy);
          this._safeRelease(sheet);
        }
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Build metadata for a closed workbook path (opened read-only in the background).
   * @param {{ path: string, sheetName?: string }} options
   * @returns {{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object, llmContext?: object }}
   */
  getClosedWorkbookMetadata(options = {}) {
    if (!options.path) {
      return { success: false, message: 'Missing workbook path.' };
    }

    try {
      return this._withWorkbookAtPath(options.path, (workbook, excel) => {
        let sheetsProxy = null;
        let sheet = null;
        if (options.sheetName) {
          sheetsProxy = workbook.Sheets;
          sheet = sheetsProxy.Item(options.sheetName);
        } else {
          sheet = workbook.ActiveSheet;
        }
        try {
          const metadata = this._collectWorksheetMetadata({
            excel,
            workbook,
            sheet,
            includeSelection: false
          });
          return {
            success: true,
            message: `Metadata read for "${sheet.Name}"`,
            ...metadata
          };
        } finally {
          this._safeRelease(sheet, sheetsProxy);
        }
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Fetch active workbook + modules + procedures + shortcut audit in one COM session.
   * This is the fast path used by renderer search hydration.
   * @returns {{
   *   success: boolean,
   *   workbook: { name: string, path: string, activeSheet: string, sheets: string[] } | null,
   *   modules: Array,
   *   procedures: Array,
   *   shortcutAudit: { success: boolean, shortcuts: Array, unmapped: Array, note?: string, message?: string },
   *   message?: string
   * }}
   */
  getActiveWorkbookContext(options = {}) {
    const { activate = true } = options;
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        const workbookInfo = this._describeWorkbookWithActiveSheet(workbook);
        const vbProject = this._getVBProjectForWorkbook(workbook);
        try {
          const modules = this._listModulesForWorkbook(workbook, vbProject);
          const procedures = this._listProceduresForWorkbook(workbook, vbProject);

          let shortcutAudit = {
            success: true,
            shortcuts: [],
            unmapped: [],
            note: 'Excel does not expose global shortcut listings. Only MacroFlow-tracked shortcuts are available.'
          };
          try {
            shortcutAudit = this._auditShortcutsForWorkbook(workbook);
          } catch (error) {
            shortcutAudit = {
              success: false,
              shortcuts: [],
              unmapped: [],
              message: String(error?.message || 'Shortcut audit failed.')
            };
          }

          return {
            success: true,
            workbook: workbookInfo,
            modules,
            procedures,
            shortcutAudit
          };
        } finally {
          this._safeRelease(vbProject);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbook: null,
        modules: [],
        procedures: [],
        shortcutAudit: { success: false, shortcuts: [], unmapped: [] },
        message: error.message
      };
    }
  }

  /**
   * Fetch open workbook list and all workbook modules in one COM session.
   * This avoids per-workbook round trips for the workbook picker and "All Files".
   * @returns {{
   *   success: boolean,
   *   workbooks: Array<{ name: string, path: string }>,
   *   allFilesModules: Array,
   *   message?: string
   * }}
   */
  getOpenWorkbookListContext(options = {}) {
    const { activate = true } = options;
    try {
      return this._withExcelApp((excel) => {
        const workbooks = [];
        const allFilesModules = [];
        const workbookRefs = [];
        const vbProjectRefs = [];
        let workbookCollection = null;

        try {
          workbookCollection = excel.Workbooks;
          const count = workbookCollection ? Number(workbookCollection.Count) : 0;
          for (let i = 1; i <= count; i += 1) {
            const workbook = workbookCollection.Item(i);
            if (!workbook) {
              continue;
            }
            workbookRefs.push(workbook);

            const workbookInfo = this._describeWorkbook(workbook);
            workbooks.push(workbookInfo);

            try {
              const vbProject = this._getVBProjectForWorkbook(workbook);
              vbProjectRefs.push(vbProject);
              const modules = this._listModulesForWorkbook(workbook, vbProject);
              modules.forEach((module) => {
                allFilesModules.push({
                  ...module,
                  workbookName: workbookInfo.name,
                  workbookPath: workbookInfo.path
                });
              });
            } catch {
              // Ignore workbook-scoped VBA failures so the list can still render.
            }
          }
        } finally {
          this._safeRelease(...vbProjectRefs, ...workbookRefs, workbookCollection);
        }

        return {
          success: true,
          workbooks,
          allFilesModules
        };
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbooks: [],
        allFilesModules: [],
        message: error.message
      };
    }
  }

  // ===========================================================================
  // CELL OPERATIONS (LEGACY)
  // ===========================================================================

  /**
   * Read a cell value.
   * @param {string} address - Cell address (e.g., 'A1', 'B2:C5')
   * @returns {{ success: boolean, value: any, address: string }}
   */
  readCell(address) {
    try {
      return this._withExcelApp((excel) => {
        let activeSheet = null;
        try {
          activeSheet = excel.ActiveSheet;
          const range = activeSheet.Range(address);
          try {
            return { success: true, address, value: this._normalizeValue(range.Value2) };
          } finally {
            this._safeRelease(range);
          }
        } finally {
          this._safeRelease(activeSheet);
        }
      });
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Write a value to a cell.
   * @param {string} address - Cell address
   * @param {any} value - Value to write
   * @returns {{ success: boolean, address: string }}
   */
  writeCell(address, value) {
    try {
      return this._withExcelApp((excel) => {
        let activeSheet = null;
        try {
          activeSheet = excel.ActiveSheet;
          const range = activeSheet.Range(address);
          try {
            range.Value2 = value;
            return { success: true, address };
          } finally {
            this._safeRelease(range);
          }
        } finally {
          this._safeRelease(activeSheet);
        }
      });
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Get the currently selected cell.
   * @returns {{ success: boolean, address: string, value: any }}
   */
  getSelection() {
    try {
      return this._withExcelApp((excel) => {
        const cell = excel.ActiveCell;
        try {
          return {
            success: true,
            address: String(cell.Address).replace(/\$/g, ''),
            value: this._normalizeValue(cell.Value2)
          };
        } finally {
          this._safeRelease(cell);
        }
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Highlight the current selection with a color.
   * @param {string} colorName - 'Yellow', 'Green', 'Red', or 'None'
   * @returns {{ success: boolean, color: string }}
   */
  highlightSelection(colorName) {
    const colorMap = {
      Yellow: 65535,
      Green: 5296274,
      Red: 255,
      None: -4142
    };

    try {
      return this._withExcelApp((excel) => {
        const selection = excel.Selection;
        let interior = null;

        try {
          if (!selection) {
            return { success: false, message: 'No cells selected' };
          }

          const colorValue = colorMap[colorName];
          if (colorValue === undefined) {
            return { success: false, message: `Unknown color: ${colorName}` };
          }

          interior = selection.Interior;
          if (colorName === 'None') {
            interior.ColorIndex = -4142;
          } else {
            interior.Color = colorValue;
          }

          return { success: true, color: colorName };
        } finally {
          this._safeRelease(interior, selection);
        }
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  // ===========================================================================
  // VBA MODULES & PROCEDURES
  // ===========================================================================

  /**
   * List modules in the active workbook's VBA project.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, modules: Array }}
   */
  listModules(options = {}) {
    const { activate = true } = options;
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        const vbProject = this._getVBProjectForWorkbook(workbook);
        try {
          const modules = this._listModulesForWorkbook(workbook, vbProject);

          return {
            success: true,
            workbook: this._describeWorkbook(workbook),
            modules
          };
        } finally {
          this._safeRelease(vbProject);
        }
      }, { activate });
    } catch (error) {
      return { success: false, modules: [], message: error.message };
    }
  }

  /**
   * List modules in a specific open workbook's VBA project.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, modules: Array, message?: string }}
   */
  listModulesByWorkbookName(workbookName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        modules: [],
        message: 'Workbook name or workbook path is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            modules: [],
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        try {
          const vbProject = this._getVBProjectForWorkbook(workbook);
          try {
            const modules = this._listModulesForWorkbook(workbook, vbProject);

            return {
              success: true,
              workbookFound: true,
              workbook: this._describeWorkbook(workbook),
              modules
            };
          } finally {
            this._safeRelease(vbProject);
          }
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        modules: [],
        message: error.message
      };
    }
  }

  /**
   * List procedures (Subs/Functions/Properties) in the active workbook.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, procedures: Array }}
   */
  listProcedures(options = {}) {
    const { activate = true } = options;
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        const vbProject = this._getVBProjectForWorkbook(workbook);
        try {
          const procedures = this._listProceduresForWorkbook(workbook, vbProject);

          return {
            success: true,
            workbook: this._describeWorkbook(workbook),
            procedures
          };
        } finally {
          this._safeRelease(vbProject);
        }
      }, { activate });
    } catch (error) {
      this._proceduresCache = null;
      return { success: false, procedures: [], message: error.message };
    }
  }

  /**
   * List procedures (Subs/Functions/Properties) for a specific open workbook name.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, procedures: Array, message?: string }}
   */
  listProceduresByWorkbookName(workbookName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        procedures: [],
        message: 'Workbook name or workbook path is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            procedures: [],
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        try {
          const vbProject = this._getVBProjectForWorkbook(workbook);
          try {
            const procedures = this._listProceduresForWorkbook(workbook, vbProject);

            return {
              success: true,
              workbookFound: true,
              workbook: this._describeWorkbook(workbook),
              procedures
            };
          } finally {
            this._safeRelease(vbProject);
          }
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      this._proceduresCache = null;
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        procedures: [],
        message: error.message
      };
    }
  }

  /**
   * Read module code for a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {{ activate?: boolean, workbookPath?: string }} options
   * @returns {{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: { name: string, path: string } | null, moduleName?: string, lineCount?: number, hash?: string, code?: string, message?: string }}
   */
  getModuleCodeByWorkbookName(workbookName, moduleName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    const normalizedModuleName = String(moduleName || '').trim();

    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: normalizedModuleName,
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        code: '',
        message: 'Workbook name or workbook path is required.'
      };
    }

    if (!normalizedModuleName) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: '',
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        code: '',
        message: 'Module name is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });

        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            moduleFound: false,
            workbook: null,
            moduleName: normalizedModuleName,
            lineCount: 0,
            hash: this._hashFnv1aHex(''),
            code: '',
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        const workbookSummary = this._describeWorkbook(workbook);
        try {
          const vbProject = this._getVBProjectForWorkbook(workbook);
          try {
            const result = this._getModuleCodeForWorkbook(vbProject, normalizedModuleName);
            return {
              ...result,
              workbookFound: true,
              workbook: workbookSummary
            };
          } finally {
            this._safeRelease(vbProject);
          }
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: normalizedModuleName,
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        code: '',
        message: error.message
      };
    }
  }

  /**
   * Read module signature (line count + hash) for a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {{ activate?: boolean, workbookPath?: string }} options
   * @returns {{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: { name: string, path: string } | null, moduleName?: string, lineCount?: number, hash?: string, message?: string }}
   */
  getModuleSignatureByWorkbookName(workbookName, moduleName, options = {}) {
    const result = this.getModuleCodeByWorkbookName(workbookName, moduleName, options);
    if (Object.prototype.hasOwnProperty.call(result, 'code')) {
      delete result.code;
    }
    return result;
  }

  /**
   * Set module code for a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {string} code
   * @param {{ activate?: boolean, workbookPath?: string, createIfMissing?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: { name: string, path: string } | null, moduleName?: string, lineCount?: number, hash?: string, message?: string }}
   */
  setModuleCodeByWorkbookName(workbookName, moduleName, code, options = {}) {
    const { activate = true, workbookPath = '', createIfMissing = false } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    const normalizedModuleName = String(moduleName || '').trim();

    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: normalizedModuleName,
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        message: 'Workbook name or workbook path is required.'
      };
    }

    if (!normalizedModuleName) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: '',
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        message: 'Module name is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            moduleFound: false,
            workbook: null,
            moduleName: normalizedModuleName,
            lineCount: 0,
            hash: this._hashFnv1aHex(''),
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        const workbookSummary = this._describeWorkbook(workbook);
        try {
          const vbProject = this._getVBProjectForWorkbook(workbook);
          try {
            const result = this._setModuleCodeForWorkbook(vbProject, normalizedModuleName, code, {
              createIfMissing
            });

            return {
              ...result,
              workbookFound: true,
              workbook: workbookSummary
            };
          } finally {
            this._safeRelease(vbProject);
          }
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: normalizedModuleName,
        lineCount: 0,
        hash: this._hashFnv1aHex(''),
        message: error.message
      };
    }
  }

  /**
   * Rename a standard VBA module in a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {string} nextModuleName
   * @param {{ activate?: boolean, workbookPath?: string }} options
   * @returns {{
   *   success: boolean,
   *   workbookFound: boolean,
   *   moduleFound: boolean,
   *   renamed: boolean,
   *   workbook?: { name: string, path: string } | null,
   *   previousModuleName?: string,
   *   moduleName?: string,
   *   message?: string
   * }}
   */
  renameModuleByWorkbookName(workbookName, moduleName, nextModuleName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedWorkbookName = String(workbookName || '').trim();
    const normalizedWorkbookPath = String(workbookPath || '').trim();
    const normalizedModuleName = String(moduleName || '').trim();
    const normalizedNextModuleName = String(nextModuleName || '').trim();

    if (!normalizedWorkbookName && !normalizedWorkbookPath) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        renamed: false,
        workbook: null,
        previousModuleName: normalizedModuleName,
        moduleName: normalizedNextModuleName,
        message: 'Workbook name or workbook path is required.'
      };
    }

    if (!normalizedModuleName) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        renamed: false,
        workbook: null,
        previousModuleName: '',
        moduleName: normalizedNextModuleName,
        message: 'Module name is required.'
      };
    }

    if (!normalizedNextModuleName) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        renamed: false,
        workbook: null,
        previousModuleName: normalizedModuleName,
        moduleName: '',
        message: 'New module name is required.'
      };
    }

    if (!this._isValidVbaModuleName(normalizedNextModuleName)) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        renamed: false,
        workbook: null,
        previousModuleName: normalizedModuleName,
        moduleName: normalizedNextModuleName,
        message: 'Invalid module name. Use letters, numbers, and underscores, and start with a letter.'
      };
    }

    const sourceNameLower = normalizedModuleName.toLowerCase();
    const targetNameLower = normalizedNextModuleName.toLowerCase();

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedWorkbookName,
          workbookPath: normalizedWorkbookPath
        });

        if (!workbook) {
          const workbookLabel = normalizedWorkbookPath || normalizedWorkbookName;
          return {
            success: true,
            workbookFound: false,
            moduleFound: false,
            renamed: false,
            workbook: null,
            previousModuleName: normalizedModuleName,
            moduleName: normalizedNextModuleName,
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        const workbookSummary = this._describeWorkbook(workbook);
        if (sourceNameLower === targetNameLower) {
          return {
            success: true,
            workbookFound: true,
            moduleFound: true,
            renamed: false,
            workbook: workbookSummary,
            previousModuleName: normalizedModuleName,
            moduleName: normalizedModuleName,
            message: 'Module name is unchanged.'
          };
        }
        let vbProject = null;
        let sourceComponent = null;
        let targetComponent = null;
        try {
          vbProject = this._getVBProjectForWorkbook(workbook);
          sourceComponent = this._findComponentByName(vbProject, normalizedModuleName);
          if (!sourceComponent) {
            return {
              success: true,
              workbookFound: true,
              moduleFound: false,
              renamed: false,
              workbook: workbookSummary,
              previousModuleName: normalizedModuleName,
              moduleName: normalizedNextModuleName,
              message: `Module "${normalizedModuleName}" was not found.`
            };
          }

          if (!this._isStandardModuleComponent(sourceComponent)) {
            return {
              success: false,
              workbookFound: true,
              moduleFound: true,
              renamed: false,
              workbook: workbookSummary,
              previousModuleName: normalizedModuleName,
              moduleName: normalizedNextModuleName,
              message: `Only standard modules can be renamed. "${normalizedModuleName}" is not a standard module.`
            };
          }

          targetComponent = this._findComponentByName(vbProject, normalizedNextModuleName);
          if (targetComponent) {
            return {
              success: false,
              workbookFound: true,
              moduleFound: true,
              renamed: false,
              workbook: workbookSummary,
              previousModuleName: normalizedModuleName,
              moduleName: normalizedNextModuleName,
              message: `Module "${normalizedNextModuleName}" already exists.`
            };
          }

          sourceComponent.Name = normalizedNextModuleName;
          return {
            success: true,
            workbookFound: true,
            moduleFound: true,
            renamed: true,
            workbook: workbookSummary,
            previousModuleName: normalizedModuleName,
            moduleName: normalizedNextModuleName,
            message: `Renamed module "${normalizedModuleName}" to "${normalizedNextModuleName}".`
          };
        } finally {
          this._safeRelease(targetComponent, sourceComponent, vbProject, workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        renamed: false,
        workbook: null,
        previousModuleName: normalizedModuleName,
        moduleName: normalizedNextModuleName,
        message: error.message
      };
    }
  }

  /**
   * Delete a standard VBA module in a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {{ activate?: boolean, workbookPath?: string }} options
   * @returns {{
   *   success: boolean,
   *   workbookFound: boolean,
   *   moduleFound: boolean,
   *   deleted: boolean,
   *   workbook?: { name: string, path: string } | null,
   *   moduleName?: string,
   *   message?: string
   * }}
   */
  deleteModuleByWorkbookName(workbookName, moduleName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedWorkbookName = String(workbookName || '').trim();
    const normalizedWorkbookPath = String(workbookPath || '').trim();
    const normalizedModuleName = String(moduleName || '').trim();

    if (!normalizedWorkbookName && !normalizedWorkbookPath) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        deleted: false,
        workbook: null,
        moduleName: normalizedModuleName,
        message: 'Workbook name or workbook path is required.'
      };
    }

    if (!normalizedModuleName) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        deleted: false,
        workbook: null,
        moduleName: '',
        message: 'Module name is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedWorkbookName,
          workbookPath: normalizedWorkbookPath
        });

        if (!workbook) {
          const workbookLabel = normalizedWorkbookPath || normalizedWorkbookName;
          return {
            success: true,
            workbookFound: false,
            moduleFound: false,
            deleted: false,
            workbook: null,
            moduleName: normalizedModuleName,
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        const workbookSummary = this._describeWorkbook(workbook);
        let vbProject = null;
        let sourceComponent = null;
        let componentsProxy = null;
        try {
          vbProject = this._getVBProjectForWorkbook(workbook);
          sourceComponent = this._findComponentByName(vbProject, normalizedModuleName);
          if (!sourceComponent) {
            return {
              success: true,
              workbookFound: true,
              moduleFound: false,
              deleted: false,
              workbook: workbookSummary,
              moduleName: normalizedModuleName,
              message: `Module "${normalizedModuleName}" was not found.`
            };
          }

          if (!this._isStandardModuleComponent(sourceComponent)) {
            return {
              success: false,
              workbookFound: true,
              moduleFound: true,
              deleted: false,
              workbook: workbookSummary,
              moduleName: normalizedModuleName,
              message: `Only standard modules can be deleted. "${normalizedModuleName}" is not a standard module.`
            };
          }

          componentsProxy = vbProject.VBComponents;
          componentsProxy.Remove(sourceComponent);
          return {
            success: true,
            workbookFound: true,
            moduleFound: true,
            deleted: true,
            workbook: workbookSummary,
            moduleName: normalizedModuleName,
            message: `Deleted module "${normalizedModuleName}".`
          };
        } finally {
          this._safeRelease(sourceComponent, componentsProxy, vbProject, workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        moduleFound: false,
        deleted: false,
        workbook: null,
        moduleName: normalizedModuleName,
        message: error.message
      };
    }
  }

  /**
   * Inject VBA code into a standard module for a specific workbook.
   * Creates the module when missing (if enabled) and replaces existing code.
   * @param {object} vbProject - Workbook VBProject COM object
   * @param {string} moduleName
   * @param {string} code
   * @param {{ createIfMissing?: boolean }} options
   * @returns {{ success: boolean, moduleFound?: boolean, created?: boolean, moduleName?: string, message?: string }}
   * @private
   */
  _injectModuleForWorkbook(vbProject, moduleName, code, options = {}) {
    const { createIfMissing = true } = options;
    const normalizedModuleName = String(moduleName || '').trim();
    if (!normalizedModuleName) {
      return { success: false, message: 'Module name is required.' };
    }

    const normalizedCode = String(code || '');
    const existing = this._findComponentByName(vbProject, normalizedModuleName);
    const hasExisting = Boolean(existing);
    this._safeRelease(existing);

    if (!hasExisting && !createIfMissing) {
      return {
        success: false,
        moduleFound: false,
        moduleName: normalizedModuleName,
        message: `Module "${normalizedModuleName}" was not found.`
      };
    }

    if (hasExisting) {
      this._removeModule(vbProject, normalizedModuleName);
    }

    let newModule = null;
    let codeModule = null;
    let componentsProxy = null;
    try {
      componentsProxy = vbProject.VBComponents;
      newModule = componentsProxy.Add(VBA_COMPONENT_TYPE.STANDARD_MODULE);
      newModule.Name = normalizedModuleName;
      codeModule = newModule.CodeModule;
      codeModule.InsertLines(codeModule.CountOfLines + 1, normalizedCode);
    } finally {
      this._safeRelease(codeModule, newModule, componentsProxy);
    }

    return {
      success: true,
      moduleFound: hasExisting,
      created: !hasExisting,
      moduleName: normalizedModuleName
    };
  }

  /**
   * Inject or replace a VBA module in the active workbook.
   * @param {string} moduleName - Name of the module (e.g., 'MacroFlowModule')
   * @param {string} code - VBA code to inject
   * @returns {{ success: boolean, message: string }}
   */
  injectModule(moduleName, code) {
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        const vbProject = this._getVBProjectForWorkbook(workbook);
        try {
          const result = this._injectModuleForWorkbook(vbProject, moduleName, code, {
            createIfMissing: true
          });
          if (!result.success) {
            return result;
          }

          const action = result.created ? 'created' : 'updated';
          return {
            success: true,
            moduleName: result.moduleName,
            message: `Module "${result.moduleName}" ${action} successfully`
          };
        } finally {
          this._safeRelease(vbProject);
        }
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Inject or replace a VBA module in a specific open workbook.
   * @param {string} workbookName
   * @param {string} moduleName
   * @param {string} code
   * @param {{ activate?: boolean, workbookPath?: string, createIfMissing?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, moduleName?: string, message: string }}
   */
  injectModuleByWorkbookName(workbookName, moduleName, code, options = {}) {
    const { activate = true, workbookPath = '', createIfMissing = true } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: 'Workbook name or workbook path is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        const workbookSummary = this._describeWorkbook(workbook);
        try {
          const vbProject = this._getVBProjectForWorkbook(workbook);
          try {
            const result = this._injectModuleForWorkbook(vbProject, moduleName, code, {
              createIfMissing
            });
            if (!result.success) {
              return {
                ...result,
                workbookFound: true,
                workbook: workbookSummary
              };
            }

            const action = result.created ? 'created' : 'updated';
            return {
              success: true,
              workbookFound: true,
              workbook: workbookSummary,
              moduleName: result.moduleName,
              message: `Module "${result.moduleName}" ${action} successfully`
            };
          } finally {
            this._safeRelease(vbProject);
          }
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: error.message
      };
    }
  }

  /**
   * Remove a module from the VBA project.
   * @param {object} vbProject - VBProject COM object
   * @param {string} moduleName - Name of module to remove
   * @private
   */
  _removeModule(vbProject, moduleName) {
    let components = null;
    try {
      components = vbProject.VBComponents;
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        if (component.Name === moduleName) {
          try {
            components.Remove(component);
          } finally {
            this._safeRelease(component);
          }
          return;
        }
        this._safeRelease(component);
      }
    } catch (error) {
      // Module doesn't exist or can't be removed.
    } finally {
      this._safeRelease(components);
    }
  }

  /**
   * Run a VBA macro by name.
   * @param {string} macroName - e.g., 'ModuleName.MacroName' or 'MacroName'
   * @returns {{ success: boolean, message: string }}
   */
  runMacro(macroName) {
    return this.runMacroWithTrap(macroName);
  }

  // ===========================================================================
  // RUNTIME + SHORTCUTS
  // ===========================================================================

  /**
   * Run a macro with a VBA error trap so runtime errors can be returned to Electron.
   * @param {string} macroName
   * @returns {{ success: boolean, message: string, error?: { number: number, description: string, source: string } }}
   */
  runMacroWithTrap(macroName) {
    try {
      return this._withActiveWorkbook(({ excel, workbook }) => {
        const addinRuntimeTarget = this._buildAddinRuntimeTarget();
        const workbookRuntimeTarget = this._buildWorkbookRuntimeTarget(workbook);
        const targetMacro = this._normalizeMacroName(macroName, workbook);

        if (!targetMacro) {
          return { success: false, message: 'Macro name is required.' };
        }

        const prep = this._prepareMacroRun(excel, workbook);
        if (!prep.ready) {
          return { success: false, message: prep.message };
        }

        logger.debug('[ExcelBridge] macro-run runtime attempt', { source: 'addin' });
        try {
          const addinResult = excel.Run(addinRuntimeTarget, targetMacro);
          const parsedAddinError = this._parseRuntimeTrapResult(addinResult);
          if (parsedAddinError) {
            return parsedAddinError;
          }

          logger.debug('[ExcelBridge] macro-run runtime source resolved', { source: 'addin' });
          return { success: true, message: `Executed "${targetMacro}"` };
        } catch (addinError) {
          if (!this._isRuntimeUnavailableError(addinError)) {
            throw addinError;
          }

          logger.debug('[ExcelBridge] macro-run runtime fallback', {
            from: 'addin',
            to: 'workbook',
            reason: String(addinError?.message || addinError || 'runtime unavailable')
          });

          try {
            this._ensureRuntimeModule(workbook);
            logger.debug('[ExcelBridge] macro-run runtime attempt', { source: 'workbook' });
            const workbookResult = excel.Run(workbookRuntimeTarget, targetMacro);
            const parsedWorkbookError = this._parseRuntimeTrapResult(workbookResult);
            if (parsedWorkbookError) {
              return {
                ...parsedWorkbookError,
                error: {
                  ...parsedWorkbookError.error,
                  runtimeSourceTried: ['addin', 'workbook'],
                  fallbackUsed: true
                }
              };
            }

            logger.debug('[ExcelBridge] macro-run runtime source resolved', { source: 'workbook' });
            return { success: true, message: `Executed "${targetMacro}"` };
          } catch (fallbackError) {
            const fallbackMessage = String(fallbackError?.message || fallbackError || 'Unknown fallback error');
            return {
              success: false,
              message: `Failed to run macro: ${fallbackMessage}`,
              error: {
                kind: 'runtime',
                rawMessage: fallbackMessage,
                runtimeSourceTried: ['addin', 'workbook'],
                fallbackUsed: true
              }
            };
          }
        }
      });
    } catch (error) {
      const rawMessage = error && error.message ? String(error.message) : 'Unknown error';
      const hresultMatch = rawMessage.match(/0x[0-9a-fA-F]+/);
      const hresult = hresultMatch ? hresultMatch[0] : '';
      const isExcelRunFailure = rawMessage.includes('0x800a9c68');
      const message = isExcelRunFailure
        ? 'Excel failed before VBA executed (likely compile error, missing reference, or dialog).'
        : `Failed to run macro: ${rawMessage}`;

      return {
        success: false,
        message,
        error: {
          kind: isExcelRunFailure ? 'excel-run-failed' : 'runtime',
          hresult,
          rawMessage
        }
      };
    }
  }

  _loadShortcutRegistry(workbook) {
    const registryName = 'MacroFlow_Shortcuts';
    let props = null;
    let prop = null;

    try {
      props = workbook.CustomDocumentProperties;

      try {
        prop = props.Item(registryName);
      } catch (error) {
        prop = null;
      }

      if (!prop) {
        props.Add(registryName, false, 4, '{}');
        prop = props.Item(registryName);
      }

      try {
        return JSON.parse(String(prop.Value || '{}'));
      } catch (error) {
        return {};
      }
    } finally {
      this._safeRelease(prop, props);
    }
  }

  _saveShortcutRegistry(workbook, registry) {
    const registryName = 'MacroFlow_Shortcuts';
    let props = null;
    let prop = null;

    try {
      props = workbook.CustomDocumentProperties;

      try {
        prop = props.Item(registryName);
      } catch (error) {
        prop = null;
      }

      if (!prop) {
        props.Add(registryName, false, 4, JSON.stringify(registry));
        return;
      }

      prop.Value = JSON.stringify(registry);
    } finally {
      this._safeRelease(prop, props);
    }
  }

  _listMacroProceduresForWorkbook(workbook) {
    try {
      const vbProject = this._getVBProjectForWorkbook(workbook);
      try {
        const procedures = this._listProceduresForWorkbook(workbook, vbProject);
        return procedures
          .filter((proc) => String(proc?.kind || '').startsWith('Sub'))
          .map((proc) => ({
            name: proc.name,
            module: proc.module
          }));
      } finally {
        this._safeRelease(vbProject);
      }
    } catch (error) {
      this._proceduresCache = null;
      return [];
    }
  }

  _listMacroProcedures(options = {}) {
    const { activate = true } = options;
    try {
      return this._withActiveWorkbook(({ workbook }) => {
        const procedures = this._listMacroProceduresForWorkbook(workbook);
        return { procedures };
      }, { activate });
    } catch (error) {
      this._proceduresCache = null;
      return { procedures: [] };
    }
  }

  _setMacroShortcutForWorkbook(excel, workbook, macroName, shortcutKey) {
    const procedures = this._listMacroProceduresForWorkbook(workbook);
    const resolvedName = this._normalizeMacroNameForWorkbook(macroName, workbook, procedures);
    excel.MacroOptions(resolvedName || macroName, null, null, null, null, shortcutKey, null, null, null, null);

    const registry = this._loadShortcutRegistry(workbook);
    if (macroName) {
      registry[macroName] = shortcutKey;
    }
    if (resolvedName) {
      registry[resolvedName] = shortcutKey;
    }
    this._saveShortcutRegistry(workbook, registry);

    return { success: true, message: `Shortcut set for ${macroName}` };
  }

  /**
   * Set a macro shortcut and track it in a workbook registry.
   * @param {string} macroName
   * @param {string} shortcutKey
   * @returns {{ success: boolean, message: string }}
   */
  setMacroShortcut(macroName, shortcutKey) {
    try {
      return this._withActiveWorkbook(({ excel, workbook }) =>
        this._setMacroShortcutForWorkbook(excel, workbook, macroName, shortcutKey)
      );
    } catch (error) {
      return { success: false, message: `Failed to set shortcut: ${error.message}` };
    }
  }

  /**
   * Set a macro shortcut in a specific open workbook and track it in that workbook registry.
   * @param {string} workbookName
   * @param {string} macroName
   * @param {string} shortcutKey
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, message: string }}
   */
  setMacroShortcutByWorkbookName(workbookName, macroName, shortcutKey, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: 'Workbook name or workbook path is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        try {
          const result = this._setMacroShortcutForWorkbook(excel, workbook, macroName, shortcutKey);
          return {
            ...result,
            workbookFound: true,
            workbook: this._describeWorkbook(workbook)
          };
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: `Failed to set shortcut: ${error.message}`
      };
    }
  }

  _auditShortcutsForWorkbook(workbook) {
    const procedures = this._listMacroProceduresForWorkbook(workbook);
    const registry = this._loadShortcutRegistry(workbook);
    const shortcuts = [];
    const unmapped = [];
    const workbookName = String(workbook?.Name || '').trim();
    const qualifiedWorkbookName = this._qualifyWorkbookName(workbookName);

    procedures.forEach((proc) => {
      const fullName = `${proc.module}.${proc.name}`;
      const qualifiedNameRaw = workbookName ? `${workbookName}!${fullName}` : fullName;
      const qualifiedNameEscaped = qualifiedWorkbookName
        ? `${qualifiedWorkbookName}!${fullName}`
        : qualifiedNameRaw;
      const shortcut =
        registry[qualifiedNameEscaped] ||
        registry[qualifiedNameRaw] ||
        registry[fullName] ||
        registry[proc.name] ||
        null;
      if (shortcut) {
        shortcuts.push({
          macro: qualifiedNameEscaped,
          shortcut
        });
      } else {
        unmapped.push(qualifiedNameEscaped);
      }
    });

    return {
      success: true,
      shortcuts,
      unmapped,
      note: 'Excel does not expose global shortcut listings. Only MacroFlow-tracked shortcuts are available.'
    };
  }

  /**
   * Audit known shortcuts for macros in the active workbook.
   * Note: Excel does not expose a full shortcut map, so only tracked shortcuts are returned.
   * @returns {{ success: boolean, shortcuts: Array, unmapped: Array, note: string }}
   */
  auditShortcuts() {
    try {
      return this._withActiveWorkbook(({ workbook }) => this._auditShortcutsForWorkbook(workbook));
    } catch (error) {
      return { success: false, shortcuts: [], unmapped: [], message: error.message };
    }
  }

  /**
   * Audit known shortcuts for macros in a specific open workbook.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, shortcuts: Array, unmapped: Array, note?: string, message?: string }}
   */
  auditShortcutsByWorkbookName(workbookName, options = {}) {
    const { activate = true, workbookPath = '' } = options;
    const normalizedName = String(workbookName || '').trim();
    const normalizedPath = String(workbookPath || '').trim();
    if (!normalizedName && !normalizedPath) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        shortcuts: [],
        unmapped: [],
        message: 'Workbook name or workbook path is required.'
      };
    }

    try {
      return this._withExcelApp((excel) => {
        const workbook = this._findOpenWorkbook(excel, {
          workbookName: normalizedName,
          workbookPath: normalizedPath
        });
        if (!workbook) {
          const workbookLabel = normalizedPath || normalizedName;
          return {
            success: true,
            workbookFound: false,
            workbook: null,
            shortcuts: [],
            unmapped: [],
            message: `Workbook "${workbookLabel}" is not open.`
          };
        }

        try {
          const result = this._auditShortcutsForWorkbook(workbook);
          return {
            ...result,
            workbookFound: true,
            workbook: this._describeWorkbook(workbook)
          };
        } finally {
          this._safeRelease(workbook);
        }
      }, { activate });
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        shortcuts: [],
        unmapped: [],
        message: error.message
      };
    }
  }
}

module.exports = new ExcelBridge();

