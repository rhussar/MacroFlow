/**
 * IPC Handlers - Routes frontend requests to ExcelBridge
 *
 * This file is a thin routing layer. All Excel logic lives in excel-bridge.js.
 * Each handler simply:
 *   1. Receives the IPC call from the renderer
 *   2. Calls the appropriate ExcelBridge method
 *   3. Returns the result
 *
 * To add a new feature:
 *   1. Add the method to ExcelBridge
 *   2. Add the IPC handler here
 *   3. Expose it in preload.js
 */

const { ipcMain, app, BrowserWindow, shell } = require('electron');
const excel = require('./excel-bridge');
const { generateVba, generateVbaStream } = require('./llm-client');
const localAiManager = require('./local-ai-manager');
const {
  loadSecurityPolicy,
  isModuleAllowed,
  isMacroAllowed,
  normalizeMacroTargetForPolicy
} = require('./security-policy');
const { initializeAuditLog, writeAuditEvent } = require('./audit-log');
const {
  logger,
  checkExcelModalState,
  collectDiagnostics,
  checkAddinStatus,
  checkRibbonStatus
} = require('./diagnostics');

// Configuration for focus handling
const FOCUS_CONFIG = {
  // Base delay before restoring alwaysOnTop (ms)
  baseRestoreDelay: 150,
  // Extra delay if Excel might be showing a modal (ms)  
  modalExtraDelay: 500,
  // Maximum number of modal check attempts
  maxModalCheckAttempts: 3,
  // Delay between modal check attempts (ms)
  modalCheckInterval: 200
};

let aiStatusEventsRegistered = false;

function sendToAllRenderers(channel, payload) {
  const windows = BrowserWindow.getAllWindows();
  for (const win of windows) {
    if (win?.webContents && !win.webContents.isDestroyed()) {
      win.webContents.send(channel, payload);
    }
  }
}

/**
 * Simple logger for IPC events (uses diagnostics logger)
 * @param {string} channel - IPC channel name
 * @param {string} phase - 'start' | 'end' | 'error'
 * @param {object} [details] - Additional details to log
 */
function logIpc(channel, phase, details = {}) {
  logger.debug('IPC', `${channel} ${phase}`, details);
}

function toSafeString(value) {
  return String(value || '').trim();
}

function normalizeWorkbookIdentity(workbookName, workbookPath) {
  return {
    name: toSafeString(workbookName),
    path: toSafeString(workbookPath)
  };
}

function hasOnlyKeys(payload, allowedKeys) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return false;
  }

  const allowedSet = new Set(allowedKeys);
  return Object.keys(payload).every((key) => allowedSet.has(key));
}

function toLowerPath(pathValue) {
  return toSafeString(pathValue).toLowerCase();
}

function toLowerName(nameValue) {
  return toSafeString(nameValue).toLowerCase();
}

function workbooksMatchScope(scope, target) {
  const scopePath = toLowerPath(scope?.path);
  const targetPath = toLowerPath(target?.path);

  if (scopePath && targetPath) {
    return scopePath === targetPath;
  }

  const scopeName = toLowerName(scope?.name);
  const targetName = toLowerName(target?.name);
  if (scopeName && targetName) {
    return scopeName === targetName;
  }

  return false;
}

function extractWorkbookNameFromMacroTarget(macroName) {
  const raw = toSafeString(macroName);
  const bangIndex = raw.indexOf('!');
  if (bangIndex < 1) {
    return '';
  }

  const workbookPrefix = raw.slice(0, bangIndex).trim();
  if (workbookPrefix.startsWith("'") && workbookPrefix.endsWith("'")) {
    return workbookPrefix.slice(1, -1).replace(/''/g, "'").trim();
  }

  return workbookPrefix;
}

function buildBlockedResult(reasonCode, message) {
  return {
    success: false,
    reasonCode,
    message
  };
}

/**
 * Get the main window (for alwaysOnTop control)
 * @returns {BrowserWindow|null}
 */
function getMainWindow() {
  const windows = BrowserWindow.getAllWindows();
  return windows.length > 0 ? windows[0] : null;
}

/**
 * Wait until Excel is no longer showing a modal, or timeout
 * @param {number} maxAttempts - Maximum check attempts
 * @param {number} interval - Delay between checks (ms)
 * @returns {Promise<boolean>} - true if Excel is ready, false if still blocked
 */
async function waitForExcelReady(maxAttempts = FOCUS_CONFIG.maxModalCheckAttempts, interval = FOCUS_CONFIG.modalCheckInterval) {
  for (let i = 0; i < maxAttempts; i++) {
    const modalState = await checkExcelModalState();
    if (!modalState.hasModal) {
      return true;
    }
    logger.debug('Excel', `Waiting for Excel to be ready (attempt ${i + 1}/${maxAttempts})`, modalState);
    await new Promise(resolve => setTimeout(resolve, interval));
  }
  return false;
}

/**
 * Temporarily disable alwaysOnTop, run a function, then restore it.
 * This prevents the Electron window from hiding Excel modal dialogs (MsgBox, etc.)
 * 
 * Enhanced version with:
 * - Excel modal state detection
 * - Configurable delays
 * - Better logging
 * 
 * @param {Function} fn - Function to execute (can be async)
 * @param {object} [options] - Options
 * @param {boolean} [options.checkModal=true] - Whether to check for Excel modals before restoring
 * @param {number} [options.restoreDelay] - Custom restore delay (ms)
 * @returns {Promise<any>} - Result of the function
 */
async function withExcelFocus(fn, options = {}) {
  const {
    checkModal = true,
    restoreDelay = FOCUS_CONFIG.baseRestoreDelay
  } = options;

  const win = getMainWindow();
  const helperManagedTopmost = Boolean(win && win.__macroflowHelperManagedTopmost === true);
  let wasOnTop = false;
  const operationStart = Date.now();

  // Step 1: Disable alwaysOnTop if it's enabled
  if (win && !win.isDestroyed()) {
    wasOnTop = win.isAlwaysOnTop();
    if (wasOnTop && !helperManagedTopmost) {
      win.setAlwaysOnTop(false);
      logger.debug('Window', 'Temporarily disabled alwaysOnTop for Excel operation');
    }
  }

  try {
    // Step 2: Run the Excel operation
    return await Promise.resolve(fn());
  } finally {
    // Step 3: Restore alwaysOnTop (always runs, even if fn throws)
    if (win && !win.isDestroyed() && wasOnTop && !helperManagedTopmost) {
      const operationDuration = Date.now() - operationStart;

      // Determine the appropriate restore delay
      let finalDelay = restoreDelay;

      // For very quick operations, we might not need much delay
      // For longer operations, Excel might still be processing
      if (operationDuration > 1000) {
        finalDelay = Math.max(restoreDelay, FOCUS_CONFIG.modalExtraDelay);
      }

      // Restore with delay, optionally checking for modals first
      setTimeout(async () => {
        if (win && !win.isDestroyed()) {
          // Optionally wait for Excel to finish showing any modals
          if (checkModal) {
            const isReady = await waitForExcelReady();
            if (!isReady) {
              logger.warn('Window', 'Excel may still be showing a modal, restoring alwaysOnTop anyway');
            }
          }

          win.setAlwaysOnTop(true);
          logger.debug('Window', 'Restored alwaysOnTop', { delayMs: finalDelay });
        }
      }, finalDelay);
    } else if (helperManagedTopmost) {
      logger.debug('Window', 'Skipped alwaysOnTop restore (managed by focus helper)');
    }
  }
}

function registerHandlers() {
  if (!aiStatusEventsRegistered) {
    aiStatusEventsRegistered = true;
    localAiManager.subscribe((status) => {
      sendToAllRenderers('ai:status', status);
    });
    // Pre-populate AI status cache in the background.
    localAiManager.getStatus().catch(() => {});
  }

  try {
    const auditInit = initializeAuditLog();
    logger.info('Audit', 'initialized', {
      auditDirectoryPath: auditInit?.auditDirectoryPath
    });
  } catch (error) {
    logger.warn('Audit', 'initialization failed', { error: error?.message || String(error) });
  }

  let isAppQuitting = false;
  let closeRequested = false;
  let activeExcelOperations = 0;
  let excelOperationSequence = 0;
  let selectedWorkbookScope = { name: '', path: '' };
  let lastSelectedWorkbookLogKey = '';

  const getSelectedWorkbookScope = () => selectedWorkbookScope;

  const setSelectedWorkbookScope = (args = {}) => {
    const normalized = normalizeWorkbookIdentity(args?.workbookName, args?.workbookPath);
    selectedWorkbookScope = normalized;
    return normalized;
  };

  const getWorkbookScopeLogKey = (scope = {}) => {
    return `${toLowerPath(scope?.path)}::${toLowerName(scope?.name)}`;
  };

  const hasSelectedWorkbookScope = () => {
    const scope = getSelectedWorkbookScope();
    return Boolean(scope.name || scope.path);
  };

  const getPolicy = () => loadSecurityPolicy();

  const writeHighRiskAudit = async (entry) => {
    try {
      const auditResult = await writeAuditEvent(entry);
      logger.info('Audit', 'high-risk action', {
        action: entry?.action,
        channel: entry?.channel,
        outcome: entry?.outcome,
        reasonCode: entry?.reasonCode,
        filePath: auditResult?.filePath
      });
    } catch (error) {
      logger.warn('Audit', 'write failed', { error: error?.message || String(error) });
    }
  };

  const getWorkbookScopeBlock = (targetWorkbook) => {
    // PERSONAL.XLSB is a global macro store – always in-scope regardless of
    // which workbook the user has selected in the picker.
    const targetName = toLowerName(targetWorkbook?.name);
    if (targetName === 'personal.xlsb') {
      return null;
    }

    if (!hasSelectedWorkbookScope()) {
      return buildBlockedResult(
        'WORKBOOK_SCOPE_NOT_SET',
        'A selected workbook is required before high-risk actions can run.'
      );
    }

    const scope = getSelectedWorkbookScope();
    if (!workbooksMatchScope(scope, targetWorkbook)) {
      return buildBlockedResult(
        'WORKBOOK_SCOPE_MISMATCH',
        'High-risk action blocked because workbook is outside the selected workbook boundary.'
      );
    }

    return null;
  };

  const getPolicyUnavailableBlock = (policyResult) => buildBlockedResult(
    'POLICY_UNAVAILABLE',
    toSafeString(policyResult?.message) || 'Security policy is unavailable.'
  );

  // Hard shutdown latch: once quit starts, never touch Excel COM again.
  app.on('before-quit', () => {
    isAppQuitting = true;
    const excelProcessIds = excel.getExcelProcessIds?.() || [];
    logger.info('Lifecycle', 'before-quit received; enabling Excel shutdown latch');
    logger.info('Lifecycle', 'Excel process snapshot at before-quit', {
      excelProcessIds,
      excelProcessCount: excelProcessIds.length
    });
    try {
      excel.setShuttingDown?.(true);
    } catch {
      // Ignore shutdown flag errors.
    }
  });

  const buildShutdownResult = () => ({
    success: false,
    message: 'APP_SHUTTING_DOWN: MacroFlow is closing and Excel operations are paused.'
  });

  const waitForExcelOperationsToDrain = (timeoutMs = 2000, pollMs = 50) => new Promise((resolve) => {
    if (activeExcelOperations < 1) {
      resolve({ drained: true, remaining: 0, waitedMs: 0 });
      return;
    }

    const startedAt = Date.now();
    const timer = setInterval(() => {
      const waitedMs = Date.now() - startedAt;
      if (activeExcelOperations < 1) {
        clearInterval(timer);
        resolve({ drained: true, remaining: 0, waitedMs });
        return;
      }
      if (waitedMs >= timeoutMs) {
        clearInterval(timer);
        resolve({
          drained: false,
          remaining: activeExcelOperations,
          waitedMs
        });
      }
    }, pollMs);
  });

  // Cooldown to prevent rapid re-attachment to a dying Excel process.
  let lastNoExcelAt = 0;
  const NO_EXCEL_COOLDOWN_MS = 3000;

  const withComRelease = async (operation, context = {}) => {
    if (isAppQuitting) {
      logger.warn('IPC', 'Excel operation blocked because app is shutting down');
      return buildShutdownResult();
    }

    // If we recently got a connection-state error, skip COM operations briefly
    // to avoid re-attaching to a dying process and creating ghost instances.
    const elapsed = Date.now() - lastNoExcelAt;
    if (lastNoExcelAt > 0 && elapsed < NO_EXCEL_COOLDOWN_MS) {
      return { success: false, message: 'NO_EXCEL: Waiting for Excel to restart.' };
    }

    const normalizedChannel = toSafeString(context?.channel) || 'unknown';
    const providedOpId = Number(context?.opId);
    const opId = Number.isFinite(providedOpId) && providedOpId > 0
      ? providedOpId
      : (++excelOperationSequence);
    const operationContext = {
      channel: normalizedChannel,
      opId
    };

    activeExcelOperations += 1;
    try {
      const result = await (
        typeof excel.runWithOperationContext === 'function'
          ? excel.runWithOperationContext(operationContext, () => Promise.resolve(operation()))
          : Promise.resolve(operation())
      );
      // Successful operation clears the cooldown.
      if (result && result.success !== false) {
        lastNoExcelAt = 0;
      }

      return result;
    } catch (err) {
      const errorMessage = String(err?.message || '');
      if (errorMessage.includes('NO_EXCEL') || errorMessage.includes('NO_VISIBLE_WINDOWS')) {
        lastNoExcelAt = Date.now();
        // Force immediate GC to release any transient COM proxies from
        // the failed operation before they can keep a ghost Excel alive.
        if (typeof global.gc === 'function') {
          try { global.gc(); } catch { /* best-effort */ }
        }
      }
      throw err;
    } finally {
      activeExcelOperations = Math.max(0, activeExcelOperations - 1);
      try {
        excel.clearComCache();
      } catch {
        // Ignore cache clear failures.
      }
    }
  };

  const withHighRiskComRelease = async (operation, context = {}) => {
    try {
      return await withComRelease(operation, context);
    } finally {
      if (typeof global.gc === 'function') {
        try { global.gc(); } catch { /* best-effort */ }
      }
    }
  };

  const withChannelComRelease = (channel, operation) => withComRelease(operation, { channel });

  const withChannelHighRiskComRelease = (channel, operation) =>
    withHighRiskComRelease(operation, { channel });

  const POLLING_PAUSE_PROTECTED_CHANNELS = new Set([
    'workbook:info',
    'workbook:list',
    'workbook:context',
    'workbook:list-context',
    'vba:modules',
    'vba:procedures',
    'vba:shortcut:audit',
    'vba:shortcut:audit:by-workbook'
  ]);

  let pollingPaused = false;
  let pollingPauseReason = '';
  let pollingPausedAt = 0;

  const extractMessage = (result) => {
    if (!result || typeof result !== 'object') {
      return '';
    }

    const message = String(result.message || result.error || '').trim();
    return message;
  };

  const extractPauseReason = (result) => {
    const message = extractMessage(result).toUpperCase();
    if (message.includes('NO_VISIBLE_WINDOWS')) {
      return 'NO_VISIBLE_WINDOWS';
    }
    if (message.includes('NO_WORKBOOK')) {
      return 'NO_WORKBOOK';
    }
    if (message.includes('NO_EXCEL')) {
      return 'NO_EXCEL';
    }
    return '';
  };

  const isPauseTrigger = (result) => {
    const reason = extractPauseReason(result);
    return reason === 'NO_EXCEL' || reason === 'NO_WORKBOOK' || reason === 'NO_VISIBLE_WINDOWS';
  };

  const setPollingPaused = (reason) => {
    const normalizedReason = reason || 'NO_EXCEL';
    const wasPaused = pollingPaused;
    const previousReason = pollingPauseReason;
    pollingPaused = true;
    pollingPauseReason = normalizedReason;
    pollingPausedAt = Date.now();
    if (!wasPaused || previousReason !== normalizedReason) {
      logger.info('IPC', 'Polling paused', {
        reason: normalizedReason,
        pausedAt: pollingPausedAt
      });
    }
  };

  const clearPollingPaused = () => {
    const wasPaused = pollingPaused;
    const previousReason = pollingPauseReason;
    const previousPausedAt = pollingPausedAt;
    pollingPaused = false;
    pollingPauseReason = '';
    pollingPausedAt = 0;
    if (wasPaused) {
      logger.info('IPC', 'Polling resumed', {
        reason: previousReason || 'UNKNOWN',
        pausedAt: previousPausedAt
      });
    }
  };

  const buildPausedResult = (channel) => {
    const code = pollingPauseReason === 'NO_WORKBOOK'
      ? 'NO_WORKBOOK'
      : (pollingPauseReason === 'NO_VISIBLE_WINDOWS' ? 'NO_VISIBLE_WINDOWS' : 'NO_EXCEL');
    const message = `${code}: Search polling is paused until reconnect succeeds.`;
    const base = {
      success: false,
      message,
      paused: true,
      reason: 'polling_paused',
      reasonCode: code,
      pausedAt: pollingPausedAt
    };

    if (channel === 'vba:modules') {
      return { ...base, modules: [] };
    }

    if (channel === 'vba:procedures') {
      return { ...base, procedures: [] };
    }

    if (channel === 'workbook:list') {
      return { ...base, workbooks: [] };
    }

    if (channel === 'workbook:context') {
      return {
        ...base,
        workbook: null,
        modules: [],
        procedures: [],
        shortcutAudit: {
          success: false,
          shortcuts: [],
          unmapped: []
        }
      };
    }

    if (channel === 'workbook:list-context') {
      return { ...base, workbooks: [], allFilesModules: [] };
    }

    if (channel === 'vba:shortcut:audit' || channel === 'vba:shortcut:audit:by-workbook') {
      return { ...base, shortcuts: [], unmapped: [] };
    }

    if (channel === 'workbook:info') {
      return {
        ...base,
        name: '',
        path: '',
        activeSheet: '',
        sheets: []
      };
    }

    return base;
  };

  const withPollingPause = async (channel, operation) => {
    const isProtected = POLLING_PAUSE_PROTECTED_CHANNELS.has(channel);
    if (isProtected && pollingPaused) {
      return buildPausedResult(channel);
    }

    const result = await withChannelComRelease(channel, operation);

    if (isProtected) {
      if (result?.success) {
        clearPollingPaused();
      } else if (isPauseTrigger(result)) {
        setPollingPaused(extractPauseReason(result));
      }
    }

    return result;
  };

  const WORKBOOK_CONTEXT_BURST_CACHE_MS = 1000;
  let workbookContextInFlight = null;
  let workbookContextCachedAt = 0;
  let workbookContextCachedResult = null;

  const clearWorkbookContextBurstCache = () => {
    workbookContextCachedAt = 0;
    workbookContextCachedResult = null;
  };

  const getWorkbookContextWithBurstCache = async () => {
    const now = Date.now();
    if (
      workbookContextCachedResult?.success &&
      now - workbookContextCachedAt < WORKBOOK_CONTEXT_BURST_CACHE_MS
    ) {
      return workbookContextCachedResult;
    }

    if (workbookContextInFlight) {
      return workbookContextInFlight;
    }

    workbookContextInFlight = (async () => {
      const result = await withPollingPause('workbook:context', () => excel.getActiveWorkbookContext());
      if (result?.success) {
        workbookContextCachedAt = Date.now();
        workbookContextCachedResult = result;
      }
      return result;
    })().finally(() => {
      workbookContextInFlight = null;
    });

    return workbookContextInFlight;
  };

  let resolveInstanceInFlight = null;
  let lastReconnectNoExcelAt = 0;
  const RECONNECT_NO_EXCEL_COOLDOWN_MS = 3000;

  const toFiniteNumberOrUndefined = (value) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  };

  const normalizeResolveMetadata = (payload) => {
    if (!payload || typeof payload !== 'object') {
      return {};
    }

    const strategy = payload.strategy === 'foreground' || payload.strategy === 'max_workbooks'
      ? payload.strategy
      : undefined;

    return {
      pid: toFiniteNumberOrUndefined(payload.pid),
      workbookCount: toFiniteNumberOrUndefined(payload.workbookCount),
      strategy
    };
  };

  const runResolveInstance = async () => {
    // Step 1: Check if we're already connected to an active workbook.
    const currentInfo = excel.getWorkbookInfo();
    if (currentInfo.success) {
      return { resolved: true, reason: 'already-connected', attempt: 0 };
    }

    // If the error isn't workbook/instance related, surface that directly.
    const msg = String(currentInfo.message || '').toUpperCase();
    if (!msg.includes('NO_WORKBOOK') && !msg.includes('MULTI_INSTANCE')) {
      return {
        resolved: false,
        reason: 'different-error',
        message: currentInfo.message
      };
    }

    // Step 2: Ask the C# helper to select and foreground the most likely Excel instance.
    const helper = excel._focusHelper;
    if (!helper || typeof helper.findExcelWithWorkbooks !== 'function') {
      return { resolved: false, reason: 'helper-unavailable' };
    }

    let helperResult;
    try {
      helperResult = await helper.findExcelWithWorkbooks(3000);
    } catch {
      return { resolved: false, reason: 'helper-error' };
    }

    const helperMeta = normalizeResolveMetadata(helperResult);

    if (!helperResult || !helperResult.found) {
      return {
        resolved: false,
        reason: String(helperResult?.reason || 'no-qualifying-instance'),
        ...helperMeta
      };
    }

    if (!helperResult.activated) {
      return { resolved: false, reason: 'activation-failed', ...helperMeta };
    }

    // Step 3: Poll-retry - ROT update after SetForegroundWindow is not instant.
    const maxRetries = 3;
    const retryDelayMs = 150;

    for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
      await new Promise((resolve) => setTimeout(resolve, retryDelayMs));
      excel.clearComCache();
      const retryInfo = excel.getWorkbookInfo();
      if (retryInfo.success) {
        return { resolved: true, reason: 'helper-resolved', attempt, ...helperMeta };
      }
    }

    return { resolved: false, reason: 'poll-exhausted', attempt: maxRetries, ...helperMeta };
  };

  const resolveInstanceOnce = async () => {
    if (!resolveInstanceInFlight) {
      resolveInstanceInFlight = withChannelComRelease('excel:resolveInstance', runResolveInstance).finally(() => {
        resolveInstanceInFlight = null;
      });
    }
    return resolveInstanceInFlight;
  };

  // ==========================================================================
  // APP CONTROLS
  // ==========================================================================
  ipcMain.handle('security:set-selected-workbook', async (_, args = {}) => {
    const allowedKeys = ['workbookName', 'workbookPath'];
    if (!hasOnlyKeys(args, allowedKeys)) {
      return buildBlockedResult(
        'VALIDATION_FAILED',
        'Invalid selected workbook payload.'
      );
    }

    const workbook = setSelectedWorkbookScope(args);
    const selected = Boolean(workbook.name || workbook.path);
    const nextLogKey = getWorkbookScopeLogKey(workbook);
    if (nextLogKey !== lastSelectedWorkbookLogKey) {
      lastSelectedWorkbookLogKey = nextLogKey;
      logger.debug('Security', 'selected workbook updated', {
        selected,
        workbookName: workbook.name,
        hasWorkbookPath: Boolean(workbook.path)
      });
    }

    return {
      success: true,
      selected,
      workbookName: workbook.name,
      workbookPath: workbook.path
    };
  });

  ipcMain.on('app:minimize', () => {
    const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
    if (win && !win.isDestroyed()) {
      win.minimize();
    }
  });

  ipcMain.on('app:open-external', (_event, url) => {
    if (typeof url === 'string' && (url.startsWith('https://') || url.startsWith('http://') || url.startsWith('mailto:'))) {
      shell.openExternal(url).catch(() => {});
    }
  });

  ipcMain.on('app:close', () => {
    if (closeRequested) {
      logger.info('Lifecycle', 'app:close ignored (close already requested)');
      return;
    }
    closeRequested = true;

    logIpc('app:close', 'start');
    isAppQuitting = true;
    const excelProcessIds = excel.getExcelProcessIds?.() || [];
    logger.info('Lifecycle', 'app:close requested; enabling Excel shutdown latch');
    logger.info('Lifecycle', 'Excel process snapshot at app:close', {
      excelProcessIds,
      excelProcessCount: excelProcessIds.length
    });
    try {
      excel.setShuttingDown?.(true);
    } catch {
      // Ignore shutdown flag errors.
    }

    void (async () => {
      const drain = await waitForExcelOperationsToDrain();
      logger.info('Lifecycle', 'Excel operations drain before quit', {
        drained: drain.drained,
        remaining: drain.remaining,
        waitedMs: drain.waitedMs
      });

      // Force V8 GC to release any lingering COM proxy wrappers.
      // Double-GC: first pass finalizes weak refs, second collects survivors.
      if (typeof global.gc === 'function') {
        try {
          global.gc();
          await new Promise((resolve) => setTimeout(resolve, 100));
          global.gc();
          logger.info('Lifecycle', 'V8 GC forced before quit');
        } catch (err) {
          logger.warn('Lifecycle', 'V8 GC force failed', { error: err.message });
        }
      }

      app.quit();
    })();
  });

  // ==========================================================================
  // VBA OPERATIONS
  // ==========================================================================

  /**
   * Inject VBA code into a module
   * Channel: 'vba:inject'
   * Args: { moduleName: string, code: string }
   */
  ipcMain.handle('vba:inject', async (_, args = {}) => {
    const moduleName = toSafeString(args?.moduleName) || 'MacroFlowModule';
    const code = String(args?.code || '');
    logIpc('vba:inject', 'start', { moduleName, codeLength: code.length });

    const blocked = buildBlockedResult(
      'VALIDATION_FAILED',
      'Use workbook-scoped injection through vba:inject:by-workbook.'
    );

    await writeHighRiskAudit({
      action: 'vba.inject',
      channel: 'vba:inject',
      outcome: 'blocked',
      reasonCode: blocked.reasonCode,
      workbook: getSelectedWorkbookScope(),
      target: { moduleName, macroName: '' },
      payload: { code, codeLength: code.length },
      result: blocked
    });

    logIpc('vba:inject', 'end', { success: false, reasonCode: blocked.reasonCode });
    return blocked;
  });

  /**
   * Inject VBA code into a module in a specific open workbook.
   * Channel: 'vba:inject:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean }
   */
  ipcMain.handle('vba:inject:by-workbook', async (_, args = {}) => {
    const {
      workbookName,
      workbookPath,
      moduleName = 'MacroFlowModule',
      code,
      createIfMissing = true
    } = args || {};

    const payload = args && typeof args === 'object' ? args : {
      workbookName,
      workbookPath,
      moduleName,
      code,
      createIfMissing
    };

    logIpc('vba:inject:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      codeLength: String(code || '').length,
      createIfMissing
    });

    const allowedKeys = ['workbookName', 'workbookPath', 'moduleName', 'code', 'createIfMissing'];
    if (!hasOnlyKeys(payload, allowedKeys)) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Invalid vba:inject:by-workbook payload.');
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizeWorkbookIdentity(workbookName, workbookPath),
        target: { moduleName: toSafeString(moduleName), macroName: '' },
        payload: { code: String(code || ''), codeLength: String(code || '').length },
        result: blocked
      });
      return blocked;
    }

    const policyResult = getPolicy();
    if (!policyResult.ok) {
      const blocked = getPolicyUnavailableBlock(policyResult);
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizeWorkbookIdentity(workbookName, workbookPath),
        target: { moduleName: toSafeString(moduleName), macroName: '' },
        payload: { code: String(code || ''), codeLength: String(code || '').length },
        result: blocked
      });
      return blocked;
    }

    const policy = policyResult.policy;
    const normalizedWorkbook = normalizeWorkbookIdentity(workbookName, workbookPath);
    const normalizedModuleName = toSafeString(moduleName);
    const normalizedCode = String(code || '');
    const normalizedCreateIfMissing = Boolean(createIfMissing);

    if ((!normalizedWorkbook.name && !normalizedWorkbook.path) || !normalizedModuleName || typeof code !== 'string') {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Workbook, module name, and code are required.');
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    if (normalizedModuleName.length > policy.limits.maxModuleNameChars || normalizedCode.length > policy.limits.maxVbaCodeChars) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Module name or code length exceeds policy limits.');
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    const scopeBlock = getWorkbookScopeBlock(normalizedWorkbook);
    if (scopeBlock) {
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: scopeBlock.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: scopeBlock
      });
      return scopeBlock;
    }

    if (policy.enforcement.denyByDefault && !isModuleAllowed(policy, normalizedModuleName)) {
      const blocked = buildBlockedResult('POLICY_DENIED', `Module "${normalizedModuleName}" is not allowlisted.`);
      await writeHighRiskAudit({
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    const result = await withChannelHighRiskComRelease('vba:inject:by-workbook',
      () => withExcelFocus(
        () => excel.injectModuleByWorkbookName(normalizedWorkbook.name, normalizedModuleName, normalizedCode, {
          workbookPath: normalizedWorkbook.path,
          createIfMissing: normalizedCreateIfMissing
        })
      )
    );

    clearWorkbookContextBurstCache();

    await writeHighRiskAudit({
      action: 'vba.inject',
      channel: 'vba:inject:by-workbook',
      outcome: result?.success ? 'succeeded' : 'failed',
      reasonCode: result?.success ? 'OK' : 'EXECUTION_FAILED',
      workbook: normalizeWorkbookIdentity(result?.workbook?.name || normalizedWorkbook.name, result?.workbook?.path || normalizedWorkbook.path),
      target: { moduleName: normalizedModuleName, macroName: '' },
      payload: { code: normalizedCode, codeLength: normalizedCode.length },
      result
    });

    logIpc('vba:inject:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleName: result.moduleName
    });
    return result;
  });

  /**
   * Read VBA module code in a specific open workbook.
   * Channel: 'vba:module-code:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string }
   */
  ipcMain.handle('vba:module-code:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = ''
  } = {}) => {
    logIpc('vba:module-code:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName
    });

    const result = await withChannelComRelease('vba:module-code:by-workbook',
      () => excel.getModuleCodeByWorkbookName(workbookName, moduleName, { workbookPath })
    );

    logIpc('vba:module-code:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      moduleName: result.moduleName,
      lineCount: result.lineCount
    });
    return result;
  });

  /**
   * Read VBA module signature in a specific open workbook.
   * Channel: 'vba:module-signature:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string }
   */
  ipcMain.handle('vba:module-signature:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = ''
  } = {}) => {
    logIpc('vba:module-signature:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName
    });

    const result = await withChannelComRelease('vba:module-signature:by-workbook',
      () => excel.getModuleSignatureByWorkbookName(workbookName, moduleName, { workbookPath })
    );

    logIpc('vba:module-signature:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      moduleName: result.moduleName,
      lineCount: result.lineCount
    });
    return result;
  });

  /**
   * Set VBA module code in a specific open workbook.
   * Channel: 'vba:module-code:set:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean }
   */
  ipcMain.handle('vba:module-code:set:by-workbook', async (_, args = {}) => {
    const {
      workbookName,
      workbookPath,
      moduleName = '',
      code = '',
      createIfMissing = false
    } = args || {};

    logIpc('vba:module-code:set:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      codeLength: code?.length,
      createIfMissing
    });

    const allowedKeys = ['workbookName', 'workbookPath', 'moduleName', 'code', 'createIfMissing'];
    if (!hasOnlyKeys(args, allowedKeys)) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Invalid vba:module-code:set:by-workbook payload.');
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizeWorkbookIdentity(workbookName, workbookPath),
        target: { moduleName: toSafeString(moduleName), macroName: '' },
        payload: { code: String(code || ''), codeLength: String(code || '').length },
        result: blocked
      });
      return blocked;
    }

    const policyResult = getPolicy();
    if (!policyResult.ok) {
      const blocked = getPolicyUnavailableBlock(policyResult);
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizeWorkbookIdentity(workbookName, workbookPath),
        target: { moduleName: toSafeString(moduleName), macroName: '' },
        payload: { code: String(code || ''), codeLength: String(code || '').length },
        result: blocked
      });
      return blocked;
    }

    const policy = policyResult.policy;
    const normalizedWorkbook = normalizeWorkbookIdentity(workbookName, workbookPath);
    const normalizedModuleName = toSafeString(moduleName);
    const normalizedCode = String(code || '');
    const normalizedCreateIfMissing = Boolean(createIfMissing);

    if ((!normalizedWorkbook.name && !normalizedWorkbook.path) || !normalizedModuleName || typeof code !== 'string') {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Workbook, module name, and code are required.');
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    if (normalizedModuleName.length > policy.limits.maxModuleNameChars || normalizedCode.length > policy.limits.maxVbaCodeChars) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Module name or code length exceeds policy limits.');
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    const scopeBlock = getWorkbookScopeBlock(normalizedWorkbook);
    if (scopeBlock) {
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: scopeBlock.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: scopeBlock
      });
      return scopeBlock;
    }

    if (policy.enforcement.denyByDefault && !isModuleAllowed(policy, normalizedModuleName)) {
      const blocked = buildBlockedResult('POLICY_DENIED', `Module "${normalizedModuleName}" is not allowlisted.`);
      await writeHighRiskAudit({
        action: 'vba.set_module_code',
        channel: 'vba:module-code:set:by-workbook',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: normalizedWorkbook,
        target: { moduleName: normalizedModuleName, macroName: '' },
        payload: { code: normalizedCode, codeLength: normalizedCode.length },
        result: blocked
      });
      return blocked;
    }

    const result = await withChannelHighRiskComRelease('vba:module-code:set:by-workbook',
      () => excel.setModuleCodeByWorkbookName(normalizedWorkbook.name, normalizedModuleName, normalizedCode, {
        workbookPath: normalizedWorkbook.path,
        createIfMissing: normalizedCreateIfMissing
      })
    );

    clearWorkbookContextBurstCache();

    await writeHighRiskAudit({
      action: 'vba.set_module_code',
      channel: 'vba:module-code:set:by-workbook',
      outcome: result?.success ? 'succeeded' : 'failed',
      reasonCode: result?.success ? 'OK' : 'EXECUTION_FAILED',
      workbook: normalizeWorkbookIdentity(result?.workbook?.name || normalizedWorkbook.name, result?.workbook?.path || normalizedWorkbook.path),
      target: { moduleName: normalizedModuleName, macroName: '' },
      payload: { code: normalizedCode, codeLength: normalizedCode.length },
      result
    });

    logIpc('vba:module-code:set:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      moduleName: result.moduleName,
      lineCount: result.lineCount
    });
    return result;
  });

  /**
   * Rename a VBA module in a specific open workbook.
   * Channel: 'vba:module:rename:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string, nextModuleName: string }
   */
  ipcMain.handle('vba:module:rename:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = '',
    nextModuleName = ''
  } = {}) => {
    logIpc('vba:module:rename:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      nextModuleName
    });

    const result = await withChannelHighRiskComRelease('vba:module:rename:by-workbook',
      () => excel.renameModuleByWorkbookName(workbookName, moduleName, nextModuleName, { workbookPath })
    );

    if (result?.success && result?.renamed) {
      clearWorkbookContextBurstCache();
    }
    logIpc('vba:module:rename:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      renamed: result.renamed,
      moduleName: result.moduleName
    });
    return result;
  });

  /**
   * Rename a VBA macro (Sub/Function) inside a module by editing its source code.
   * Channel: 'vba:macro:rename:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string, macroName: string, nextMacroName: string }
   */
  ipcMain.handle('vba:macro:rename:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = '',
    macroName = '',
    nextMacroName = ''
  } = {}) => {
    logIpc('vba:macro:rename:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      macroName,
      nextMacroName
    });

    const result = await withChannelHighRiskComRelease('vba:macro:rename:by-workbook',
      () => excel.renameMacroByWorkbookName(workbookName, moduleName, macroName, nextMacroName, { workbookPath })
    );

    if (result?.success && result?.renamed) {
      clearWorkbookContextBurstCache();
    }
    logIpc('vba:macro:rename:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      macroFound: result.macroFound,
      renamed: result.renamed,
      macroName: result.macroName
    });
    return result;
  });

  /**
   * Delete a VBA module in a specific open workbook.
   * Channel: 'vba:module:delete:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string }
   */
  ipcMain.handle('vba:module:delete:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = ''
  } = {}) => {
    logIpc('vba:module:delete:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName
    });

    const result = await withChannelHighRiskComRelease('vba:module:delete:by-workbook',
      () => excel.deleteModuleByWorkbookName(workbookName, moduleName, { workbookPath })
    );

    if (result?.success && result?.deleted) {
      clearWorkbookContextBurstCache();
    }
    logIpc('vba:module:delete:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      moduleFound: result.moduleFound,
      deleted: result.deleted,
      moduleName: result.moduleName
    });
    return result;
  });

  /**
   * Run a VBA macro
   * Channel: 'vba:run'
   * Args: { macroName: string }
   * 
   * IMPORTANT: This operation may block if the macro shows a MsgBox or other dialog.
   * We disable alwaysOnTop so the user can see and dismiss Excel prompts.
   */
  ipcMain.handle('vba:run', async (_, args = {}) => {
    const { macroName = '' } = args || {};
    logIpc('vba:run', 'start', { macroName });

    const allowedKeys = ['macroName'];
    if (!hasOnlyKeys(args, allowedKeys)) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Invalid vba:run payload.');
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: getSelectedWorkbookScope(),
        target: { moduleName: '', macroName: toSafeString(macroName) },
        payload: { codeLength: 0 },
        result: blocked
      });
      return blocked;
    }

    const policyResult = getPolicy();
    if (!policyResult.ok) {
      const blocked = getPolicyUnavailableBlock(policyResult);
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: getSelectedWorkbookScope(),
        target: { moduleName: '', macroName: toSafeString(macroName) },
        payload: { codeLength: 0 },
        result: blocked
      });
      return blocked;
    }

    const policy = policyResult.policy;
    const normalizedMacroName = toSafeString(macroName);
    if (!normalizedMacroName) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Macro name is required.');
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: getSelectedWorkbookScope(),
        target: { moduleName: '', macroName: normalizedMacroName },
        payload: { codeLength: 0 },
        result: blocked
      });
      return blocked;
    }

    if (normalizedMacroName.length > policy.limits.maxMacroNameChars) {
      const blocked = buildBlockedResult('VALIDATION_FAILED', 'Macro name exceeds policy limits.');
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: getSelectedWorkbookScope(),
        target: { moduleName: '', macroName: normalizedMacroName },
        payload: { codeLength: 0 },
        result: blocked
      });
      return blocked;
    }

    const workbookNameFromTarget = extractWorkbookNameFromMacroTarget(normalizedMacroName);
    const targetWorkbook = workbookNameFromTarget
      ? normalizeWorkbookIdentity(workbookNameFromTarget, '')
      : getSelectedWorkbookScope();
    const scopeBlock = getWorkbookScopeBlock(targetWorkbook);
    if (scopeBlock) {
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: scopeBlock.reasonCode,
        workbook: targetWorkbook,
        target: { moduleName: '', macroName: normalizedMacroName },
        payload: { codeLength: 0 },
        result: scopeBlock
      });
      return scopeBlock;
    }

    const macroPolicyTarget = normalizeMacroTargetForPolicy(normalizedMacroName);
    if (policy.enforcement.denyByDefault && !isMacroAllowed(policy, macroPolicyTarget)) {
      const blocked = buildBlockedResult('POLICY_DENIED', `Macro "${macroPolicyTarget}" is not allowlisted.`);
      await writeHighRiskAudit({
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: blocked.reasonCode,
        workbook: targetWorkbook,
        target: { moduleName: '', macroName: normalizedMacroName },
        payload: { codeLength: 0 },
        result: blocked
      });
      return blocked;
    }

    // Use withExcelFocus - CRITICAL for MsgBox/dialog visibility
    const result = await withChannelHighRiskComRelease('vba:run',
      () => withExcelFocus(() => excel.runMacro(normalizedMacroName))
    );

    clearWorkbookContextBurstCache();
    await writeHighRiskAudit({
      action: 'vba.run',
      channel: 'vba:run',
      outcome: result?.success ? 'succeeded' : 'failed',
      reasonCode: result?.success ? 'OK' : 'EXECUTION_FAILED',
      workbook: targetWorkbook,
      target: { moduleName: '', macroName: normalizedMacroName },
      payload: { codeLength: 0 },
      result
    });

    logIpc('vba:run', 'end', { success: result.success, message: result.message });
    return result;
  });

  /**
   * Get local AI runtime/model status.
   * Channel: 'ai:status'
   */
  ipcMain.handle('ai:status', async () => {
    logIpc('ai:status', 'start');
    const result = await localAiManager.getStatus();
    logIpc('ai:status', 'end', {
      ready: result.ready,
      stage: result.stage,
      setupInProgress: result.setupInProgress
    });
    return result;
  });

  /**
   * Ensure the local AI runtime is started and ready.
   * Channel: 'ai:ensure-ready'
   */
  ipcMain.handle('ai:ensure-ready', async () => {
    logIpc('ai:ensure-ready', 'start');
    const result = await localAiManager.ensureReady();
    logIpc('ai:ensure-ready', 'end', {
      ready: result.ready,
      stage: result.stage
    });
    return result;
  });

  /**
   * Start local AI setup.
   * Channel: 'ai:setup'
   */
  ipcMain.handle('ai:setup', async () => {
    logIpc('ai:setup', 'start');
    const result = await localAiManager.setup();
    logIpc('ai:setup', 'end', {
      success: result.success,
      started: result.started,
      stage: result.status?.stage
    });
    return result;
  });

  /**
   * Remove the configured local AI model.
   * Channel: 'ai:remove-model'
   */
  ipcMain.handle('ai:remove-model', async () => {
    logIpc('ai:remove-model', 'start');
    const result = await localAiManager.removeModel();
    logIpc('ai:remove-model', 'end', {
      success: result.success,
      started: result.started,
      stage: result.status?.stage
    });
    return result;
  });

  /**
   * Generate VBA code from natural language prompt (local AI).
   * Channel: 'ai:generate-vba'
   * Args: { prompt: string, intent?: string, workbookName?: string, workbookPath?: string, moduleName?: string, sheetName?: string, currentCode?: string, includeCurrentCode?: boolean }
   */
  ipcMain.handle('ai:generate-vba', async (event, {
    prompt = '',
    intent = '',
    workbookName = '',
    workbookPath = '',
    moduleName = '',
    sheetName = '',
    currentCode = '',
    includeCurrentCode = false
  } = {}) => {
    const shouldIncludeCurrentCode = Boolean(includeCurrentCode);
    logIpc('ai:generate-vba', 'start', {
      intent,
      workbookName,
      workbookPath,
      moduleName,
      sheetName,
      promptChars: String(prompt || '').length,
      includeCurrentCode: shouldIncludeCurrentCode,
      currentCodeChars: shouldIncludeCurrentCode ? String(currentCode || '').length : 0
    });

    const normalizedIntent = String(intent || '').trim().toLowerCase();
    const useStreaming = normalizedIntent === 'ask';
    const params = {
      prompt,
      intent,
      workbookName,
      workbookPath,
      moduleName,
      sheetName,
      currentCode: shouldIncludeCurrentCode ? currentCode : '',
      includeCurrentCode: shouldIncludeCurrentCode
    };

    const result = useStreaming
      ? await generateVbaStream(params, {}, (token) => {
          if (event.sender && !event.sender.isDestroyed()) {
            event.sender.send('ai:generate-token', token);
          }
        })
      : await generateVba(params);

    logIpc('ai:generate-vba', 'end', {
      success: result.success,
      reason: result.reason,
      model: result.model
    });
    return result;
  });


  /**
   * List VBA modules in the active workbook
   * Channel: 'vba:modules'
   */
  ipcMain.handle('vba:modules', async () => {
    logIpc('vba:modules', 'start');

    // Read-only listing path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = await withPollingPause('vba:modules', () => excel.listModules());

    logIpc('vba:modules', 'end', { success: result.success, count: result.modules?.length });
    return result;
  });

  /**
   * List VBA modules in a specific open workbook.
   * Channel: 'vba:modules:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:modules:by-workbook', async (_, { workbookName, workbookPath } = {}) => {
    logIpc('vba:modules:by-workbook', 'start', { workbookName, workbookPath });

    const result = await withChannelComRelease('vba:modules:by-workbook',
      () => excel.listModulesByWorkbookName(workbookName, { workbookPath })
    );

    logIpc('vba:modules:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      count: result.modules?.length
    });
    return result;
  });

  /**
   * List procedures (Subs/Functions/Properties) in the active workbook
   * Channel: 'vba:procedures'
   */
  ipcMain.handle('vba:procedures', async () => {
    logIpc('vba:procedures', 'start');

    // Read-only listing path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = await withPollingPause('vba:procedures', () => excel.listProcedures());

    logIpc('vba:procedures', 'end', { success: result.success, count: result.procedures?.length });
    return result;
  });

  /**
   * List procedures (Subs/Functions/Properties) in a specific open workbook.
   * Channel: 'vba:procedures:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:procedures:by-workbook', async (_, { workbookName, workbookPath } = {}) => {
    logIpc('vba:procedures:by-workbook', 'start', { workbookName, workbookPath });

    const result = await withChannelComRelease('vba:procedures:by-workbook',
      () => excel.listProceduresByWorkbookName(workbookName, { workbookPath })
    );

    logIpc('vba:procedures:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      count: result.procedures?.length
    });
    return result;
  });

  /**
   * Set a macro shortcut and track it
   * Channel: 'vba:shortcut:set'
   * Args: { macroName: string, shortcutKey: string }
   */
  ipcMain.handle('vba:shortcut:set', async (_, { macroName, shortcutKey }) => {
    logIpc('vba:shortcut:set', 'start', { macroName, shortcutKey });

    const result = await withChannelComRelease('vba:shortcut:set',
      () => withExcelFocus(() => excel.setMacroShortcut(macroName, shortcutKey))
    );

    clearWorkbookContextBurstCache();
    logIpc('vba:shortcut:set', 'end', { success: result.success });
    return result;
  });

  /**
   * Set a macro shortcut and track it in a specific open workbook.
   * Channel: 'vba:shortcut:set:by-workbook'
   * Args: { workbookName: string, macroName: string, shortcutKey: string }
   */
  ipcMain.handle('vba:shortcut:set:by-workbook', async (_, { workbookName, workbookPath, macroName, shortcutKey } = {}) => {
    logIpc('vba:shortcut:set:by-workbook', 'start', { workbookName, workbookPath, macroName, shortcutKey });

    const result = await withChannelComRelease('vba:shortcut:set:by-workbook',
      () => withExcelFocus(
        () => excel.setMacroShortcutByWorkbookName(workbookName, macroName, shortcutKey, { workbookPath })
      )
    );

    clearWorkbookContextBurstCache();
    logIpc('vba:shortcut:set:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound
    });
    return result;
  });

  /**
   * Audit tracked shortcuts
   * Channel: 'vba:shortcut:audit'
   */
  ipcMain.handle('vba:shortcut:audit', async () => {
    logIpc('vba:shortcut:audit', 'start');

    // Read-only audit path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = await withPollingPause('vba:shortcut:audit', () => excel.auditShortcuts());

    logIpc('vba:shortcut:audit', 'end', { success: result.success });
    return result;
  });

  /**
   * Audit tracked shortcuts in a specific open workbook.
   * Channel: 'vba:shortcut:audit:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:shortcut:audit:by-workbook', async (_, { workbookName, workbookPath } = {}) => {
    logIpc('vba:shortcut:audit:by-workbook', 'start', { workbookName, workbookPath });

    const result = await withPollingPause(
      'vba:shortcut:audit:by-workbook',
      () => excel.auditShortcutsByWorkbookName(workbookName, { workbookPath })
    );

    logIpc('vba:shortcut:audit:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound
    });
    return result;
  });

  // ==========================================================================
  // CELL OPERATIONS
  // ==========================================================================

  /**
   * Read cell value
   * Channel: 'cell:read'
   * Args: { address: string }
   */
  ipcMain.handle('cell:read', async (_, { address }) => {
    logIpc('cell:read', 'start', { address });
    const result = await withChannelComRelease('cell:read', () => excel.readCell(address));
    logIpc('cell:read', 'end', { success: result.success });
    return result;
  });

  /**
   * Write cell value
   * Channel: 'cell:write'
   * Args: { address: string, value: any }
   */
  ipcMain.handle('cell:write', async (_, { address, value }) => {
    logIpc('cell:write', 'start', { address });
    const result = await withChannelComRelease('cell:write', () => excel.writeCell(address, value));
    clearWorkbookContextBurstCache();
    logIpc('cell:write', 'end', { success: result.success });
    return result;
  });

  /**
   * Get current selection
   * Channel: 'cell:selection'
   */
  ipcMain.handle('cell:selection', async () => {
    logIpc('cell:selection', 'start');
    const result = await withChannelComRelease('cell:selection', () => excel.getSelection());
    logIpc('cell:selection', 'end', { success: result.success });
    return result;
  });

  /**
   * Highlight current selection
   * Channel: 'cell:highlight'
   * Args: { color: string }
   */
  ipcMain.handle('cell:highlight', async (_, { color }) => {
    logIpc('cell:highlight', 'start', { color });
    const result = await withChannelComRelease('cell:highlight', () => excel.highlightSelection(color));
    logIpc('cell:highlight', 'end', { success: result.success });
    return result;
  });

  // ==========================================================================
  // WORKBOOK INFO
  // ==========================================================================

  /**
   * Get workbook info
   * Channel: 'workbook:info'
   */
  ipcMain.handle('workbook:info', async () => {
    logIpc('workbook:info', 'start');
    const result = await withPollingPause('workbook:info', () => excel.getWorkbookInfo());
    logIpc('workbook:info', 'end', { success: result.success, name: result.name });
    return result;
  });

  /**
   * Get all open workbooks
   * Channel: 'workbook:list'
   * Returns: { success: boolean, workbooks: Array<{ name: string, path: string }> }
   */
  ipcMain.handle('workbook:list', async () => {
    logIpc('workbook:list', 'start');
    const result = await withPollingPause('workbook:list', () => excel.getOpenWorkbooks());
    logIpc('workbook:list', 'end', { success: result.success, count: result.workbooks?.length });
    return result;
  });

  /**
   * Get active workbook context in one backend call.
   * Channel: 'workbook:context'
   * Returns: { success, workbook, modules, procedures, shortcutAudit }
   */
  ipcMain.handle('workbook:context', async () => {
    const startedAt = Date.now();
    logIpc('workbook:context', 'start');
    const result = await getWorkbookContextWithBurstCache();
    logIpc('workbook:context', 'end', {
      success: result.success,
      modules: result.modules?.length,
      procedures: result.procedures?.length,
      durationMs: Date.now() - startedAt
    });
    return result;
  });

  /**
   * Get open workbook list + all-files modules in one backend call.
   * Channel: 'workbook:list-context'
   * Returns: { success, workbooks, allFilesModules }
   */
  ipcMain.handle('workbook:list-context', async () => {
    const startedAt = Date.now();
    logIpc('workbook:list-context', 'start');
    const result = await withPollingPause(
      'workbook:list-context',
      () => excel.getOpenWorkbookListContext()
    );
    logIpc('workbook:list-context', 'end', {
      success: result.success,
      workbookCount: result.workbooks?.length,
      moduleCount: result.allFilesModules?.length,
      durationMs: Date.now() - startedAt
    });
    return result;
  });

  /**
   * Get PERSONAL.XLSB status from XLSTART + open workbook state.
   * Channel: 'personal:status'
   * Returns: { success, workbookFound, workbook, fileExists, workbookPath, message? }
   */
  ipcMain.handle('personal:status', async () => {
    logIpc('personal:status', 'start');
    const result = await withChannelComRelease('personal:status', () => excel.getPersonalWorkbookStatus());
    logIpc('personal:status', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      fileExists: result.fileExists
    });
    return result;
  });

  /**
   * Get PERSONAL.XLSB status + procedures + shortcut audit in one backend call.
   * Channel: 'personal:context'
   * Returns: { success, workbookFound, workbook, fileExists, workbookPath, procedures, shortcutAudit, message? }
   */
  ipcMain.handle('personal:context', async (_, args) => {
    const payload = args && typeof args === 'object' ? args : {};
    if (
      !hasOnlyKeys(payload, ['includeShortcutAudit']) ||
      (payload.includeShortcutAudit !== undefined && typeof payload.includeShortcutAudit !== 'boolean')
    ) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        fileExists: false,
        workbookPath: '',
        windowVisible: null,
        windowHidden: false,
        procedures: [],
        shortcutAudit: null,
        message: 'VALIDATION_FAILED: includeShortcutAudit must be a boolean when provided.'
      };
    }

    const includeShortcutAudit = payload.includeShortcutAudit !== false;
    logIpc('personal:context', 'start', { includeShortcutAudit });
    const result = await withChannelComRelease(
      'personal:context',
      () => excel.getPersonalWorkbookContext({ includeShortcutAudit })
    );
    logIpc('personal:context', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      procedures: result.procedures?.length,
      includeShortcutAudit
    });
    return result;
  });

  /**
   * Open PERSONAL.XLSB from XLSTART.
   * Channel: 'personal:open'
   * Args: { visible?: boolean }
   * Returns: { success, workbookFound, opened, alreadyOpen, workbook, fileExists, workbookPath, windowVisible, windowHidden, message? }
   */
  ipcMain.handle('personal:open', async (_, args) => {
    const payload = args && typeof args === 'object' ? args : {};
    if (
      !hasOnlyKeys(payload, ['visible']) ||
      (Object.prototype.hasOwnProperty.call(payload, 'visible') && typeof payload.visible !== 'boolean')
    ) {
      return buildBlockedResult('VALIDATION_FAILED', 'personal:open only accepts an optional boolean "visible" flag.');
    }

    logIpc('personal:open', 'start', { visible: payload.visible });
    const result = await withChannelComRelease('personal:open', () => excel.openPersonalWorkbook(payload));
    if (result?.success) {
      clearPollingPaused();
      clearWorkbookContextBurstCache();
    }
    logIpc('personal:open', 'end', {
      success: result.success,
      opened: result.opened,
      alreadyOpen: result.alreadyOpen,
      windowVisible: result.windowVisible
    });
    return result;
  });

  /**
   * Create PERSONAL.XLSB in XLSTART, then open it.
   * Channel: 'personal:create'
   * Args: { visible?: boolean }
   * Returns: { success, created, opened, workbookFound, workbook, fileExists, workbookPath, windowVisible, windowHidden, message? }
   */
  ipcMain.handle('personal:create', async (_, args) => {
    const payload = args && typeof args === 'object' ? args : {};
    if (
      !hasOnlyKeys(payload, ['visible']) ||
      (Object.prototype.hasOwnProperty.call(payload, 'visible') && typeof payload.visible !== 'boolean')
    ) {
      return buildBlockedResult('VALIDATION_FAILED', 'personal:create only accepts an optional boolean "visible" flag.');
    }

    logIpc('personal:create', 'start', { visible: payload.visible });
    const result = await withChannelComRelease('personal:create', () => excel.createPersonalWorkbook(payload));
    if (result?.success) {
      clearPollingPaused();
      clearWorkbookContextBurstCache();
    }
    logIpc('personal:create', 'end', {
      success: result.success,
      created: result.created,
      opened: result.opened,
      windowVisible: result.windowVisible
    });
    return result;
  });

  /**
   * Show or hide the open PERSONAL.XLSB workbook window.
   * Channel: 'personal:visibility:set'
   * Args: { visible: boolean }
   * Returns: { success, workbookFound, workbook, fileExists, workbookPath, windowVisible, windowHidden, visibilityChanged, message? }
   */
  ipcMain.handle('personal:visibility:set', async (_, args) => {
    const payload = args && typeof args === 'object' ? args : {};
    if (!hasOnlyKeys(payload, ['visible']) || typeof payload.visible !== 'boolean') {
      return buildBlockedResult('VALIDATION_FAILED', 'personal:visibility:set requires a boolean "visible" flag.');
    }

    logIpc('personal:visibility:set', 'start', { visible: payload.visible });
    const result = await withChannelComRelease('personal:visibility:set', () => excel.setPersonalWorkbookVisibility(payload));
    logIpc('personal:visibility:set', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      windowVisible: result.windowVisible,
      visibilityChanged: result.visibilityChanged
    });
    return result;
  });

  /**
   * Open the Excel XLSTART folder for PERSONAL.XLSB.
   * Channel: 'personal:open-folder'
   * Returns: { success, workbookPath, folderPath, fileExists, message? }
   */
  ipcMain.handle('personal:open-folder', async () => {
    logIpc('personal:open-folder', 'start');
    const location = excel.getPersonalWorkbookLocation();
    if (!location?.success) {
      logIpc('personal:open-folder', 'end', { success: false });
      return location;
    }

    try {
      if (location.fileExists && location.workbookPath) {
        shell.showItemInFolder(location.workbookPath);
      } else {
        const openResult = await shell.openPath(location.folderPath);
        if (openResult) {
          throw new Error(String(openResult));
        }
      }

      const result = {
        success: true,
        workbookPath: location.workbookPath,
        folderPath: location.folderPath,
        fileExists: location.fileExists
      };
      logIpc('personal:open-folder', 'end', { success: true, fileExists: location.fileExists });
      return result;
    } catch (error) {
      const result = {
        success: false,
        workbookPath: location.workbookPath,
        folderPath: location.folderPath,
        fileExists: location.fileExists,
        message: String(error?.message || 'Unable to open the XLSTART folder.')
      };
      logIpc('personal:open-folder', 'end', { success: false, message: result.message });
      return result;
    }
  });

  /**
   * List worksheets with UsedRange stats
   * Channel: 'workbook:sheets'
   */
  ipcMain.handle('workbook:sheets', async () => {
    logIpc('workbook:sheets', 'start');
    const result = await withChannelComRelease('workbook:sheets', () => excel.listWorksheets());
    logIpc('workbook:sheets', 'end', { success: result.success, count: result.sheets?.length });
    return result;
  });

  /**
   * Get worksheet metadata and preview
   * Channel: 'workbook:metadata'
   * Args: { sheetName?: string }
   */
  ipcMain.handle('workbook:metadata', async (_, args) => {
    logIpc('workbook:metadata', 'start', { sheetName: args?.sheetName });
    const result = await withChannelComRelease('workbook:metadata', () => excel.getWorksheetMetadata(args));
    logIpc('workbook:metadata', 'end', { success: result.success });
    return result;
  });

  /**
   * Get worksheet metadata for a specific open workbook
   * Channel: 'workbook:metadata:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, sheetName?: string }
   */
  ipcMain.handle('workbook:metadata:by-workbook', async (_, args) => {
    const payload = args && typeof args === 'object' ? args : {};
    if (
      !hasOnlyKeys(payload, ['workbookName', 'workbookPath', 'sheetName']) ||
      (Object.prototype.hasOwnProperty.call(payload, 'workbookName') && typeof payload.workbookName !== 'string') ||
      (Object.prototype.hasOwnProperty.call(payload, 'workbookPath') && typeof payload.workbookPath !== 'string') ||
      (Object.prototype.hasOwnProperty.call(payload, 'sheetName') && typeof payload.sheetName !== 'string')
    ) {
      return buildBlockedResult(
        'VALIDATION_FAILED',
        'workbook:metadata:by-workbook accepts optional string values for "workbookName", "workbookPath", and "sheetName".'
      );
    }

    logIpc('workbook:metadata:by-workbook', 'start', {
      workbookName: payload.workbookName,
      workbookPath: payload.workbookPath,
      sheetName: payload.sheetName
    });
    const result = await withChannelComRelease(
      'workbook:metadata:by-workbook',
      () => excel.getWorksheetMetadataByWorkbookName(payload.workbookName, payload)
    );
    logIpc('workbook:metadata:by-workbook', 'end', {
      success: result.success,
      workbookFound: result.workbookFound
    });
    return result;
  });

  /**
   * Get metadata for a closed workbook path
   * Channel: 'workbook:metadata:closed'
   * Args: { path: string, sheetName?: string }
   */
  ipcMain.handle('workbook:metadata:closed', async (_, args) => {
    logIpc('workbook:metadata:closed', 'start', { path: args?.path });
    const result = await withChannelComRelease('workbook:metadata:closed', () => excel.getClosedWorkbookMetadata(args));
    logIpc('workbook:metadata:closed', 'end', { success: result.success });
    return result;
  });

  // ==========================================================================
  // MULTI-INSTANCE RESOLUTION
  // ==========================================================================

  /**
   * Attempt to silently resolve the correct Excel instance.
   * Uses the C# helper's AccessibleObjectFromWindow to find the instance
   * with user workbooks and SetForegroundWindow it so the next COM call connects correctly.
   * Falls through gracefully if the helper is unavailable.
   * Channel: 'excel:resolveInstance'
   */
  ipcMain.handle('excel:resolveInstance', async () => {
    logIpc('excel:resolveInstance', 'start');
    try {
      const result = await resolveInstanceOnce();

      if (result?.resolved) {
        clearPollingPaused();
        clearWorkbookContextBurstCache();
      }

      logIpc('excel:resolveInstance', 'end', {
        resolved: Boolean(result?.resolved),
        reason: result?.reason,
        pid: result?.pid,
        workbookCount: result?.workbookCount,
        strategy: result?.strategy,
        attempt: result?.attempt
      });

      return result;
    } catch (error) {
      const message = String(error?.message || error || 'Unknown resolve error');
      logIpc('excel:resolveInstance', 'error', { message });
      return { resolved: false, reason: 'handler-error', message };
    }
  });

  /**
   * Clear COM cache and retry connection.
   * Used for manual/auto reconnect after user focuses the correct Excel window.
   * Channel: 'excel:reconnect'
   */
  ipcMain.handle('excel:reconnect', async () => {
    logIpc('excel:reconnect', 'start');
    const elapsedSinceNoExcel = Date.now() - lastReconnectNoExcelAt;
    if (lastReconnectNoExcelAt > 0 && elapsedSinceNoExcel < RECONNECT_NO_EXCEL_COOLDOWN_MS) {
      const cooldownResult = { success: false, message: 'NO_EXCEL: Waiting for Excel to restart.' };
      logIpc('excel:reconnect', 'end', { success: false, cooldown: true });
      return cooldownResult;
    }

    const result = await withChannelComRelease('excel:reconnect', () => {
      excel.clearComCache();
      return excel.getWorkbookInfo();
    });
    const reconnectMessage = String(result?.message || result?.error || '').toUpperCase();
    if (reconnectMessage.includes('NO_EXCEL') || reconnectMessage.includes('NO_VISIBLE_WINDOWS')) {
      lastReconnectNoExcelAt = Date.now();
    } else if (result?.success) {
      lastReconnectNoExcelAt = 0;
    }
    if (result?.success) {
      clearPollingPaused();
      clearWorkbookContextBurstCache();
    }
    logIpc('excel:reconnect', 'end', { success: result.success });
    return result;
  });

  // ==========================================================================
  // DIAGNOSTICS
  // ==========================================================================

  /**
   * Get full system diagnostics
   * Channel: 'diagnostics:collect'
   * Returns: { timestamp, system, excel, addin, ribbon, recentLogs }
   */
  ipcMain.handle('diagnostics:collect', async () => {
    logIpc('diagnostics:collect', 'start');
    try {
      const result = await collectDiagnostics();
      logIpc('diagnostics:collect', 'end', { success: true });
      return { success: true, ...result };
    } catch (error) {
      logIpc('diagnostics:collect', 'error', { error: error.message });
      return { success: false, error: error.message };
    }
  });

  /**
   * Check add-in installation status
   * Channel: 'diagnostics:addin'
   * Returns: { installed, registered, loadedInExcel, ... }
   */
  ipcMain.handle('diagnostics:addin', async () => {
    logIpc('diagnostics:addin', 'start');
    try {
      const result = await checkAddinStatus();
      logIpc('diagnostics:addin', 'end', {
        installed: result.installed,
        loaded: result.loadedInExcel
      });
      return { success: true, ...result };
    } catch (error) {
      logIpc('diagnostics:addin', 'error', { error: error.message });
      return { success: false, error: error.message };
    }
  });

  /**
   * Check ribbon status
   * Channel: 'diagnostics:ribbon'
   * Returns: { hasRibbon, ribbonErrors, ribbonInfo }
   */
  ipcMain.handle('diagnostics:ribbon', async () => {
    logIpc('diagnostics:ribbon', 'start');
    try {
      const result = await checkRibbonStatus();
      logIpc('diagnostics:ribbon', 'end', { hasRibbon: result.hasRibbon });
      return { success: true, ...result };
    } catch (error) {
      logIpc('diagnostics:ribbon', 'error', { error: error.message });
      return { success: false, error: error.message };
    }
  });

  /**
   * Check Excel modal/focus state
   * Channel: 'diagnostics:excel-state'
   * Returns: { hasModal, interactive, ready, reason }
   */
  ipcMain.handle('diagnostics:excel-state', async () => {
    logIpc('diagnostics:excel-state', 'start');
    try {
      const result = await checkExcelModalState();
      logIpc('diagnostics:excel-state', 'end', { hasModal: result.hasModal });
      return { success: true, ...result };
    } catch (error) {
      logIpc('diagnostics:excel-state', 'error', { error: error.message });
      return { success: false, error: error.message };
    }
  });
}

module.exports = { registerHandlers };
