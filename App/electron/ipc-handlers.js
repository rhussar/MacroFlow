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

  // ==========================================================================
  // APP CONTROLS
  // ==========================================================================

  ipcMain.on('app:close', () => {
    logIpc('app:close', 'start');
    app.quit();
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
    const result = await withExcelFocus(() => excel.injectModule(moduleName, code));

    logIpc('vba:inject', 'end', { success: result.success });
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
    const result = await withExcelFocus(() => excel.runMacro(macroName));

    logIpc('vba:run', 'end', { success: result.success, message: result.message });
    return result;
  });


  /**
   * List VBA modules in the active workbook
   * Channel: 'vba:modules'
   */
  ipcMain.handle('vba:modules', () => {
    logIpc('vba:modules', 'start');

    // Read-only listing path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = excel.listModules();

    logIpc('vba:modules', 'end', { success: result.success, count: result.modules?.length });
    return result;
  });

  /**
   * List VBA modules in a specific open workbook.
   * Channel: 'vba:modules:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:modules:by-workbook', (_, { workbookName } = {}) => {
    logIpc('vba:modules:by-workbook', 'start', { workbookName });

    const result = excel.listModulesByWorkbookName(workbookName);

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
  ipcMain.handle('vba:procedures', () => {
    logIpc('vba:procedures', 'start');

    // Read-only listing path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = excel.listProcedures();

    logIpc('vba:procedures', 'end', { success: result.success, count: result.procedures?.length });
    return result;
  });

  /**
   * List procedures (Subs/Functions/Properties) in a specific open workbook.
   * Channel: 'vba:procedures:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:procedures:by-workbook', (_, { workbookName } = {}) => {
    logIpc('vba:procedures:by-workbook', 'start', { workbookName });

    const result = excel.listProceduresByWorkbookName(workbookName);

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

    const result = await withExcelFocus(() => excel.setMacroShortcut(macroName, shortcutKey));

    logIpc('vba:shortcut:set', 'end', { success: result.success });
    return result;
  });

  /**
   * Set a macro shortcut and track it in a specific open workbook.
   * Channel: 'vba:shortcut:set:by-workbook'
   * Args: { workbookName: string, macroName: string, shortcutKey: string }
   */
  ipcMain.handle('vba:shortcut:set:by-workbook', async (_, { workbookName, macroName, shortcutKey } = {}) => {
    logIpc('vba:shortcut:set:by-workbook', 'start', { workbookName, macroName, shortcutKey });

    const result = await withExcelFocus(() => excel.setMacroShortcutByWorkbookName(workbookName, macroName, shortcutKey));

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
  ipcMain.handle('vba:shortcut:audit', () => {
    logIpc('vba:shortcut:audit', 'start');

    // Read-only audit path: avoid withExcelFocus to reduce z-order/focus churn.
    const result = excel.auditShortcuts();

    logIpc('vba:shortcut:audit', 'end', { success: result.success });
    return result;
  });

  /**
   * Audit tracked shortcuts in a specific open workbook.
   * Channel: 'vba:shortcut:audit:by-workbook'
   * Args: { workbookName: string }
   */
  ipcMain.handle('vba:shortcut:audit:by-workbook', (_, { workbookName } = {}) => {
    logIpc('vba:shortcut:audit:by-workbook', 'start', { workbookName });

    const result = excel.auditShortcutsByWorkbookName(workbookName);

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
  ipcMain.handle('cell:read', (_, { address }) => {
    logIpc('cell:read', 'start', { address });
    const result = excel.readCell(address);
    logIpc('cell:read', 'end', { success: result.success });
    return result;
  });

  /**
   * Write cell value
   * Channel: 'cell:write'
   * Args: { address: string, value: any }
   */
  ipcMain.handle('cell:write', (_, { address, value }) => {
    logIpc('cell:write', 'start', { address });
    const result = excel.writeCell(address, value);
    logIpc('cell:write', 'end', { success: result.success });
    return result;
  });

  /**
   * Get current selection
   * Channel: 'cell:selection'
   */
  ipcMain.handle('cell:selection', () => {
    logIpc('cell:selection', 'start');
    const result = excel.getSelection();
    logIpc('cell:selection', 'end', { success: result.success });
    return result;
  });

  /**
   * Highlight current selection
   * Channel: 'cell:highlight'
   * Args: { color: string }
   */
  ipcMain.handle('cell:highlight', (_, { color }) => {
    logIpc('cell:highlight', 'start', { color });
    const result = excel.highlightSelection(color);
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
  ipcMain.handle('workbook:info', () => {
    logIpc('workbook:info', 'start');
    const result = excel.getWorkbookInfo();
    logIpc('workbook:info', 'end', { success: result.success, name: result.name });
    return result;
  });

  /**
   * Get all open workbooks
   * Channel: 'workbook:list'
   * Returns: { success: boolean, workbooks: Array<{ name: string, path: string }> }
   */
  ipcMain.handle('workbook:list', () => {
    logIpc('workbook:list', 'start');
    const result = excel.getOpenWorkbooks();
    logIpc('workbook:list', 'end', { success: result.success, count: result.workbooks?.length });
    return result;
  });

  /**
   * List worksheets with UsedRange stats
   * Channel: 'workbook:sheets'
   */
  ipcMain.handle('workbook:sheets', () => {
    logIpc('workbook:sheets', 'start');
    const result = excel.listWorksheets();
    logIpc('workbook:sheets', 'end', { success: result.success, count: result.sheets?.length });
    return result;
  });

  /**
   * Get worksheet metadata and preview
   * Channel: 'workbook:metadata'
   * Args: { sheetName?: string }
   */
  ipcMain.handle('workbook:metadata', (_, args) => {
    logIpc('workbook:metadata', 'start', { sheetName: args?.sheetName });
    const result = excel.getWorksheetMetadata(args);
    logIpc('workbook:metadata', 'end', { success: result.success });
    return result;
  });

  /**
   * Get metadata for a closed workbook path
   * Channel: 'workbook:metadata:closed'
   * Args: { path: string, sheetName?: string }
   */
  ipcMain.handle('workbook:metadata:closed', (_, args) => {
    logIpc('workbook:metadata:closed', 'start', { path: args?.path });
    const result = excel.getClosedWorkbookMetadata(args);
    logIpc('workbook:metadata:closed', 'end', { success: result.success });
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


