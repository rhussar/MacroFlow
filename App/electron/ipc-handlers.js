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

const { ipcMain, app } = require('electron');
const excel = require('./excel-bridge');

function registerHandlers() {

  // ==========================================================================
  // APP CONTROLS
  // ==========================================================================

  ipcMain.on('app:close', () => {
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
  ipcMain.handle('vba:inject', (_, { moduleName = 'MacroFlowModule', code }) => {
    return excel.injectModule(moduleName, code);
  });

  /**
   * Run a VBA macro
   * Channel: 'vba:run'
   * Args: { macroName: string }
   */
  ipcMain.handle('vba:run', (_, { macroName }) => {
    return excel.runMacro(macroName);
  });


  /**
   * List VBA modules in the active workbook
   * Channel: 'vba:modules'
   */
  ipcMain.handle('vba:modules', () => {
    return excel.listModules();
  });

  /**
   * List procedures (Subs/Functions/Properties) in the active workbook
   * Channel: 'vba:procedures'
   */
  ipcMain.handle('vba:procedures', () => {
    return excel.listProcedures();
  });

  /**
   * Set a macro shortcut and track it
   * Channel: 'vba:shortcut:set'
   * Args: { macroName: string, shortcutKey: string }
   */
  ipcMain.handle('vba:shortcut:set', (_, { macroName, shortcutKey }) => {
    return excel.setMacroShortcut(macroName, shortcutKey);
  });

  /**
   * Audit tracked shortcuts
   * Channel: 'vba:shortcut:audit'
   */
  ipcMain.handle('vba:shortcut:audit', () => {
    return excel.auditShortcuts();
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
    return excel.readCell(address);
  });

  /**
   * Write cell value
   * Channel: 'cell:write'
   * Args: { address: string, value: any }
   */
  ipcMain.handle('cell:write', (_, { address, value }) => {
    return excel.writeCell(address, value);
  });

  /**
   * Get current selection
   * Channel: 'cell:selection'
   */
  ipcMain.handle('cell:selection', () => {
    return excel.getSelection();
  });

  /**
   * Highlight current selection
   * Channel: 'cell:highlight'
   * Args: { color: string }
   */
  ipcMain.handle('cell:highlight', (_, { color }) => {
    return excel.highlightSelection(color);
  });

  // ==========================================================================
  // WORKBOOK INFO
  // ==========================================================================

  /**
   * Get workbook info
   * Channel: 'workbook:info'
   */
  ipcMain.handle('workbook:info', () => {
    return excel.getWorkbookInfo();
  });

  /**
   * Get all open workbooks
   * Channel: 'workbook:list'
   * Returns: { success: boolean, workbooks: Array<{ name: string, path: string }> }
   */
  ipcMain.handle('workbook:list', () => {
    return excel.getOpenWorkbooks();
  });

  /**
   * List worksheets with UsedRange stats
   * Channel: 'workbook:sheets'
   */
  ipcMain.handle('workbook:sheets', () => {
    return excel.listWorksheets();
  });

  /**
   * Get worksheet metadata and preview
   * Channel: 'workbook:metadata'
   * Args: { sheetName?: string }
   */
  ipcMain.handle('workbook:metadata', (_, args) => {
    return excel.getWorksheetMetadata(args);
  });

  /**
   * Get metadata for a closed workbook path
   * Channel: 'workbook:metadata:closed'
   * Args: { path: string, sheetName?: string }
   */
  ipcMain.handle('workbook:metadata:closed', (_, args) => {
    return excel.getClosedWorkbookMetadata(args);
  });
}

module.exports = { registerHandlers };
