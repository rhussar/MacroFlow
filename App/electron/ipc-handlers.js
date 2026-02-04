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

/**
 * Simple logger for IPC events
 * @param {string} channel - IPC channel name
 * @param {string} phase - 'start' | 'end' | 'error'
 * @param {object} [details] - Additional details to log
 */
function logIpc(channel, phase, details = {}) {
  const timestamp = new Date().toISOString().substr(11, 12);
  const detailStr = Object.keys(details).length > 0 
    ? ` ${JSON.stringify(details)}` 
    : '';
  console.log(`[IPC ${timestamp}] ${channel} ${phase}${detailStr}`);
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
 * Temporarily disable alwaysOnTop, run a function, then restore it.
 * This prevents the Electron window from hiding Excel modal dialogs (MsgBox, etc.)
 * @param {Function} fn - Function to execute (can be async)
 * @returns {Promise<any>} - Result of the function
 */
async function withExcelFocus(fn) {
  const win = getMainWindow();
  let wasOnTop = false;

  // Step 1: Disable alwaysOnTop if it's enabled
  if (win && !win.isDestroyed()) {
    wasOnTop = win.isAlwaysOnTop();
    if (wasOnTop) {
      win.setAlwaysOnTop(false);
      console.log('[Window] Temporarily disabled alwaysOnTop for Excel operation');
    }
  }

  try {
    // Step 2: Run the Excel operation
    return await Promise.resolve(fn());
  } finally {
    // Step 3: Restore alwaysOnTop (always runs, even if fn throws)
    if (win && !win.isDestroyed() && wasOnTop) {
      // Small delay to ensure Excel dialog can appear before we restore
      setTimeout(() => {
        if (win && !win.isDestroyed()) {
          win.setAlwaysOnTop(true);
          console.log('[Window] Restored alwaysOnTop');
        }
      }, 100);
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
  ipcMain.handle('vba:modules', async () => {
    logIpc('vba:modules', 'start');
    
    const result = await withExcelFocus(() => excel.listModules());
    
    logIpc('vba:modules', 'end', { success: result.success, count: result.modules?.length });
    return result;
  });

  /**
   * List procedures (Subs/Functions/Properties) in the active workbook
   * Channel: 'vba:procedures'
   */
  ipcMain.handle('vba:procedures', async () => {
    logIpc('vba:procedures', 'start');
    
    const result = await withExcelFocus(() => excel.listProcedures());
    
    logIpc('vba:procedures', 'end', { success: result.success, count: result.procedures?.length });
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
   * Audit tracked shortcuts
   * Channel: 'vba:shortcut:audit'
   */
  ipcMain.handle('vba:shortcut:audit', async () => {
    logIpc('vba:shortcut:audit', 'start');
    
    const result = await withExcelFocus(() => excel.auditShortcuts());
    
    logIpc('vba:shortcut:audit', 'end', { success: result.success });
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
}

module.exports = { registerHandlers };
