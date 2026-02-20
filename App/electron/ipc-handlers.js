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

const { ipcMain, app, BrowserWindow } = require('electron');
const excel = require('./excel-bridge');
const { generateVba } = require('./openai-client');
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

/**
 * Simple logger for IPC events (uses diagnostics logger)
 * @param {string} channel - IPC channel name
 * @param {string} phase - 'start' | 'end' | 'error'
 * @param {object} [details] - Additional details to log
 */
function logIpc(channel, phase, details = {}) {
  logger.debug('IPC', `${channel} ${phase}`, details);
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
  let isAppQuitting = false;
  let closeRequested = false;
  let activeExcelOperations = 0;

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

  const withComRelease = async (operation) => {
    if (isAppQuitting) {
      logger.warn('IPC', 'Excel operation blocked because app is shutting down');
      return buildShutdownResult();
    }

    activeExcelOperations += 1;
    try {
      return await Promise.resolve(operation());
    } finally {
      activeExcelOperations = Math.max(0, activeExcelOperations - 1);
      try {
        excel.clearComCache();
      } catch {
        // Ignore cache clear failures.
      }
    }
  };

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
    return reason === 'NO_EXCEL' || reason === 'NO_WORKBOOK';
  };

  const setPollingPaused = (reason) => {
    pollingPaused = true;
    pollingPauseReason = reason || 'NO_EXCEL';
    pollingPausedAt = Date.now();
  };

  const clearPollingPaused = () => {
    pollingPaused = false;
    pollingPauseReason = '';
    pollingPausedAt = 0;
  };

  const buildPausedResult = (channel) => {
    const code = pollingPauseReason === 'NO_WORKBOOK' ? 'NO_WORKBOOK' : 'NO_EXCEL';
    const message = `${code}: Search polling is paused until reconnect succeeds.`;
    const base = {
      success: false,
      message,
      paused: true,
      reason: 'polling_paused',
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

    const result = await withComRelease(operation);

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
      resolveInstanceInFlight = withComRelease(runResolveInstance).finally(() => {
        resolveInstanceInFlight = null;
      });
    }
    return resolveInstanceInFlight;
  };

  // ==========================================================================
  // APP CONTROLS
  // ==========================================================================

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
  ipcMain.handle('vba:inject', async (_, { moduleName = 'MacroFlowModule', code }) => {
    logIpc('vba:inject', 'start', { moduleName, codeLength: code?.length });

    // Use withExcelFocus to prevent hiding Excel dialogs during injection
    const result = await withComRelease(
      () => withExcelFocus(() => excel.injectModule(moduleName, code))
    );

    clearWorkbookContextBurstCache();
    logIpc('vba:inject', 'end', { success: result.success });
    return result;
  });

  /**
   * Inject VBA code into a module in a specific open workbook.
   * Channel: 'vba:inject:by-workbook'
   * Args: { workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean }
   */
  ipcMain.handle('vba:inject:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = 'MacroFlowModule',
    code,
    createIfMissing = true
  } = {}) => {
    logIpc('vba:inject:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      codeLength: code?.length,
      createIfMissing
    });

    const result = await withComRelease(
      () => withExcelFocus(
        () => excel.injectModuleByWorkbookName(workbookName, moduleName, code, {
          workbookPath,
          createIfMissing
        })
      )
    );

    clearWorkbookContextBurstCache();
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

    const result = await withComRelease(
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

    const result = await withComRelease(
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
  ipcMain.handle('vba:module-code:set:by-workbook', async (_, {
    workbookName,
    workbookPath,
    moduleName = '',
    code = '',
    createIfMissing = false
  } = {}) => {
    logIpc('vba:module-code:set:by-workbook', 'start', {
      workbookName,
      workbookPath,
      moduleName,
      codeLength: code?.length,
      createIfMissing
    });

    const result = await withComRelease(
      () => excel.setModuleCodeByWorkbookName(workbookName, moduleName, code, {
        workbookPath,
        createIfMissing
      })
    );

    clearWorkbookContextBurstCache();
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

    const result = await withComRelease(
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

    const result = await withComRelease(
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
  ipcMain.handle('vba:run', async (_, { macroName }) => {
    logIpc('vba:run', 'start', { macroName });

    // Use withExcelFocus - CRITICAL for MsgBox/dialog visibility
    const result = await withComRelease(
      () => withExcelFocus(() => excel.runMacro(macroName))
    );

    clearWorkbookContextBurstCache();
    logIpc('vba:run', 'end', { success: result.success, message: result.message });
    return result;
  });

  /**
   * Generate VBA code from natural language prompt (OpenAI).
   * Channel: 'ai:generate-vba'
   * Args: { prompt: string, workbookName?: string, moduleName?: string, currentCode?: string }
   */
  ipcMain.handle('ai:generate-vba', async (_, {
    prompt = '',
    workbookName = '',
    moduleName = '',
    currentCode = ''
  } = {}) => {
    logIpc('ai:generate-vba', 'start', {
      workbookName,
      moduleName,
      promptChars: String(prompt || '').length,
      currentCodeChars: String(currentCode || '').length
    });

    const result = await Promise.resolve(generateVba({
      prompt,
      workbookName,
      moduleName,
      currentCode
    }));

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

    const result = await withComRelease(
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

    const result = await withComRelease(
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

    const result = await withComRelease(
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

    const result = await withComRelease(
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
    const result = await withComRelease(() => excel.readCell(address));
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
    const result = await withComRelease(() => excel.writeCell(address, value));
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
    const result = await withComRelease(() => excel.getSelection());
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
    const result = await withComRelease(() => excel.highlightSelection(color));
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
    const result = await withComRelease(() => excel.getPersonalWorkbookStatus());
    logIpc('personal:status', 'end', {
      success: result.success,
      workbookFound: result.workbookFound,
      fileExists: result.fileExists
    });
    return result;
  });

  /**
   * Open PERSONAL.XLSB from XLSTART.
   * Channel: 'personal:open'
   * Returns: { success, workbookFound, opened, alreadyOpen, workbook, fileExists, workbookPath, message? }
   */
  ipcMain.handle('personal:open', async () => {
    logIpc('personal:open', 'start');
    const result = await withComRelease(() => excel.openPersonalWorkbook());
    if (result?.success) {
      clearPollingPaused();
      clearWorkbookContextBurstCache();
    }
    logIpc('personal:open', 'end', {
      success: result.success,
      opened: result.opened,
      alreadyOpen: result.alreadyOpen
    });
    return result;
  });

  /**
   * Create PERSONAL.XLSB in XLSTART, then open it.
   * Channel: 'personal:create'
   * Returns: { success, created, opened, workbookFound, workbook, fileExists, workbookPath, message? }
   */
  ipcMain.handle('personal:create', async () => {
    logIpc('personal:create', 'start');
    const result = await withComRelease(() => excel.createPersonalWorkbook());
    if (result?.success) {
      clearPollingPaused();
      clearWorkbookContextBurstCache();
    }
    logIpc('personal:create', 'end', {
      success: result.success,
      created: result.created,
      opened: result.opened
    });
    return result;
  });

  /**
   * List worksheets with UsedRange stats
   * Channel: 'workbook:sheets'
   */
  ipcMain.handle('workbook:sheets', async () => {
    logIpc('workbook:sheets', 'start');
    const result = await withComRelease(() => excel.listWorksheets());
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
    const result = await withComRelease(() => excel.getWorksheetMetadata(args));
    logIpc('workbook:metadata', 'end', { success: result.success });
    return result;
  });

  /**
   * Get metadata for a closed workbook path
   * Channel: 'workbook:metadata:closed'
   * Args: { path: string, sheetName?: string }
   */
  ipcMain.handle('workbook:metadata:closed', async (_, args) => {
    logIpc('workbook:metadata:closed', 'start', { path: args?.path });
    const result = await withComRelease(() => excel.getClosedWorkbookMetadata(args));
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
    const result = await withComRelease(() => {
      excel.clearComCache();
      return excel.getWorkbookInfo();
    });
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


