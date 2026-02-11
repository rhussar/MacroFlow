/**
 * Preload Script - Secure bridge between React UI and Electron main process
 *
 * This exposes a clean API to the renderer process via window.excel
 * All methods return Promises with { success: boolean, ... } responses
 *
 * Usage in React:
 *   const result = await window.excel.vba.inject({ code: '...' });
 *   const cell = await window.excel.cell.read({ address: 'A1' });
 *   
 * Window control (for alwaysOnTop management):
 *   await window.excel.window.setAlwaysOnTop(false);
 *   const { value } = await window.excel.window.getAlwaysOnTop();
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('excel', {

  // ==========================================================================
  // VBA OPERATIONS
  // ==========================================================================
  vba: {
    /**
     * Inject VBA code into a module
     * @param {{ moduleName?: string, code: string }} args
     * @returns {Promise<{ success: boolean, message: string }>}
     */
    inject: (args) => ipcRenderer.invoke('vba:inject', args),

    /**
     * Run a VBA macro
     * @param {{ macroName: string }} args
     * @returns {Promise<{ success: boolean, message: string }>}
     */
    run: (args) => ipcRenderer.invoke('vba:run', args),


    /**
     * List VBA modules in the active workbook
     * @returns {Promise<{ success: boolean, workbook?: object, modules: Array }>}
     */
    modules: () => ipcRenderer.invoke('vba:modules'),

    /**
     * List procedures (Subs/Functions/Properties) in the active workbook
     * @returns {Promise<{ success: boolean, workbook?: object, procedures: Array }>}
     */
    procedures: () => ipcRenderer.invoke('vba:procedures'),

    /**
     * Set a macro shortcut and track it
     * @param {{ macroName: string, shortcutKey: string }} args
     * @returns {Promise<{ success: boolean, message: string }>}
     */
    setShortcut: (args) => ipcRenderer.invoke('vba:shortcut:set', args),

    /**
     * Audit tracked shortcuts
     * @returns {Promise<{ success: boolean, shortcuts: Array, unmapped: Array }>}
     */
    auditShortcuts: () => ipcRenderer.invoke('vba:shortcut:audit'),
  },

  // ==========================================================================
  // CELL OPERATIONS
  // ==========================================================================
  cell: {
    /**
     * Read a cell value
     * @param {{ address: string }} args
     * @returns {Promise<{ success: boolean, address: string, value: any }>}
     */
    read: (args) => ipcRenderer.invoke('cell:read', args),

    /**
     * Write a value to a cell
     * @param {{ address: string, value: any }} args
     * @returns {Promise<{ success: boolean, address: string }>}
     */
    write: (args) => ipcRenderer.invoke('cell:write', args),

    /**
     * Get the current selection
     * @returns {Promise<{ success: boolean, address: string, value: any }>}
     */
    selection: () => ipcRenderer.invoke('cell:selection'),

    /**
     * Highlight current selection
     * @param {{ color: string }} args
     * @returns {Promise<{ success: boolean, color?: string }>}
     */
    highlight: (args) => ipcRenderer.invoke('cell:highlight', args),
  },

  // ==========================================================================
  // WORKBOOK OPERATIONS
  // ==========================================================================
  workbook: {
    /**
     * Get info about the active workbook
     * @returns {Promise<{ success: boolean, name: string, path: string, sheets: string[] }>}
     */
    info: () => ipcRenderer.invoke('workbook:info'),

    /**
     * Get all open workbooks
     * @returns {Promise<{ success: boolean, workbooks: Array<{ name: string, path: string }> }>}
     */
    list: () => ipcRenderer.invoke('workbook:list'),

    /**
     * List worksheets with UsedRange stats
     * @returns {Promise<{ success: boolean, sheets: Array }>}
     */
    sheets: () => ipcRenderer.invoke('workbook:sheets'),

    /**
     * Get worksheet metadata and preview
     * @param {{ sheetName?: string }} args
     * @returns {Promise<{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object }>}
     */
    metadata: (args) => ipcRenderer.invoke('workbook:metadata', args),

    /**
     * Get metadata for a closed workbook path
     * @param {{ path: string, sheetName?: string }} args
     * @returns {Promise<{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object }>}
     */
    metadataClosed: (args) => ipcRenderer.invoke('workbook:metadata:closed', args),
  },

  // ==========================================================================
  // WINDOW CONTROLS (for alwaysOnTop management)
  // ==========================================================================
  window: {
    /**
     * Set the alwaysOnTop state of the main window
     * @param {boolean} value - Whether to keep window on top
     * @returns {Promise<{ success: boolean, previousValue?: boolean, currentValue?: boolean }>}
     */
    setAlwaysOnTop: (value) => ipcRenderer.invoke('window:setAlwaysOnTop', value),

    /**
     * Get the current alwaysOnTop state
     * @returns {Promise<{ success: boolean, value?: boolean }>}
     */
    getAlwaysOnTop: () => ipcRenderer.invoke('window:getAlwaysOnTop'),
  },

  // ==========================================================================
  // DIAGNOSTICS (for troubleshooting Excel integration)
  // ==========================================================================
  diagnostics: {
    /**
     * Collect full system diagnostics
     * @returns {Promise<{ success: boolean, timestamp?: string, system?: object, excel?: object, addin?: object, ribbon?: object }>}
     */
    collect: () => ipcRenderer.invoke('diagnostics:collect'),

    /**
     * Check add-in installation status
     * @returns {Promise<{ success: boolean, installed?: boolean, registered?: boolean, loadedInExcel?: boolean }>}
     */
    addin: () => ipcRenderer.invoke('diagnostics:addin'),

    /**
     * Check ribbon status
     * @returns {Promise<{ success: boolean, hasRibbon?: boolean, ribbonErrors?: Array, ribbonInfo?: object }>}
     */
    ribbon: () => ipcRenderer.invoke('diagnostics:ribbon'),

    /**
     * Check Excel modal/focus state
     * @returns {Promise<{ success: boolean, hasModal?: boolean, interactive?: boolean, ready?: boolean }>}
     */
    excelState: () => ipcRenderer.invoke('diagnostics:excel-state'),
  },

  // ==========================================================================
  // APP CONTROLS
  // ==========================================================================
  app: {
    /**
     * Close the application
     */
    close: () => ipcRenderer.send('app:close'),
  },
});

