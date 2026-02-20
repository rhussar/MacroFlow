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
     * Inject VBA code into a module in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: object, moduleName?: string, message: string }>}
     */
    injectByWorkbook: (args) => ipcRenderer.invoke('vba:inject:by-workbook', args),

    /**
     * Read module code from a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, code?: string, message?: string }>}
     */
    moduleCodeByWorkbook: (args) => ipcRenderer.invoke('vba:module-code:by-workbook', args),

    /**
     * Read module signature (lineCount + hash) from a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, message?: string }>}
     */
    moduleSignatureByWorkbook: (args) => ipcRenderer.invoke('vba:module-signature:by-workbook', args),

    /**
     * Set module code in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, message?: string }>}
     */
    setModuleCodeByWorkbook: (args) => ipcRenderer.invoke('vba:module-code:set:by-workbook', args),

    /**
     * Rename a VBA module in a specific open workbook.
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string, nextModuleName: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, moduleFound: boolean, renamed: boolean, workbook?: object, previousModuleName?: string, moduleName?: string, message?: string }>}
     */
    renameModuleByWorkbook: (args) => ipcRenderer.invoke('vba:module:rename:by-workbook', args),

    /**
     * Delete a VBA module in a specific open workbook.
     * @param {{ workbookName?: string, workbookPath?: string, moduleName: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, moduleFound: boolean, deleted: boolean, workbook?: object, moduleName?: string, message?: string }>}
     */
    deleteModuleByWorkbook: (args) => ipcRenderer.invoke('vba:module:delete:by-workbook', args),

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
     * List VBA modules in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: object, modules: Array, message?: string }>}
     */
    modulesByWorkbook: (args) => ipcRenderer.invoke('vba:modules:by-workbook', args),

    /**
     * List procedures (Subs/Functions/Properties) in the active workbook
     * @returns {Promise<{ success: boolean, workbook?: object, procedures: Array }>}
     */
    procedures: () => ipcRenderer.invoke('vba:procedures'),

    /**
     * List procedures (Subs/Functions/Properties) in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: object, procedures: Array, message?: string }>}
     */
    proceduresByWorkbook: (args) => ipcRenderer.invoke('vba:procedures:by-workbook', args),

    /**
     * Set a macro shortcut and track it
     * @param {{ macroName: string, shortcutKey: string }} args
     * @returns {Promise<{ success: boolean, message: string }>}
     */
    setShortcut: (args) => ipcRenderer.invoke('vba:shortcut:set', args),

    /**
     * Set a macro shortcut and track it in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string, macroName: string, shortcutKey: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: object, message: string }>}
     */
    setShortcutByWorkbook: (args) => ipcRenderer.invoke('vba:shortcut:set:by-workbook', args),

    /**
     * Audit tracked shortcuts
     * @returns {Promise<{ success: boolean, shortcuts: Array, unmapped: Array }>}
     */
    auditShortcuts: () => ipcRenderer.invoke('vba:shortcut:audit'),

    /**
     * Audit tracked shortcuts in a specific open workbook
     * @param {{ workbookName?: string, workbookPath?: string }} args
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: object, shortcuts: Array, unmapped: Array, note?: string, message?: string }>}
     */
    auditShortcutsByWorkbook: (args) => ipcRenderer.invoke('vba:shortcut:audit:by-workbook', args),
  },

  // ==========================================================================
  // AI OPERATIONS
  // ==========================================================================
  ai: {
    /**
     * Generate VBA module code using OpenAI.
     * @param {{ prompt: string, workbookName?: string, moduleName?: string, currentCode?: string }} args
     * @returns {Promise<{ success: boolean, code?: string, model?: string, usage?: { promptTokens?: number, completionTokens?: number, totalTokens?: number }, reason?: string, message?: string }>}
     */
    generateVba: (args) => ipcRenderer.invoke('ai:generate-vba', args)
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
     * @returns {Promise<{ success: boolean, name: string, path: string, activeSheet: string, sheets: string[] }>}
     */
    info: () => ipcRenderer.invoke('workbook:info'),

    /**
     * Get all open workbooks
     * @returns {Promise<{ success: boolean, workbooks: Array<{ name: string, path: string }> }>}
     */
    list: () => ipcRenderer.invoke('workbook:list'),

    /**
     * Get active workbook context (workbook + modules + procedures + shortcut audit) in one call.
     * Non-breaking fast path used by search hydration when available.
     * @returns {Promise<{
     *   success: boolean,
     *   workbook: { name: string, path: string, activeSheet: string, sheets: string[] } | null,
     *   modules: Array,
     *   procedures: Array,
     *   shortcutAudit: { success: boolean, shortcuts: Array, unmapped: Array, note?: string, message?: string },
     *   message?: string
     * }>}
     */
    context: () => ipcRenderer.invoke('workbook:context'),

    /**
     * Get open workbook list and all-files modules in one call.
     * @returns {Promise<{ success: boolean, workbooks: Array<{ name: string, path: string }>, allFilesModules: Array, message?: string }>}
     */
    listContext: () => ipcRenderer.invoke('workbook:list-context'),

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
  // PERSONAL.XLSB OPERATIONS
  // ==========================================================================
  personal: {
    /**
     * Get PERSONAL.XLSB status from XLSTART + open workbook state.
     * @returns {Promise<{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, fileExists: boolean, workbookPath?: string, message?: string }>}
     */
    status: () => ipcRenderer.invoke('personal:status'),

    /**
     * Open PERSONAL.XLSB from XLSTART.
     * @returns {Promise<{ success: boolean, workbookFound: boolean, opened: boolean, alreadyOpen: boolean, workbook?: { name: string, path: string } | null, fileExists: boolean, workbookPath?: string, message?: string }>}
     */
    open: () => ipcRenderer.invoke('personal:open'),

    /**
     * Create PERSONAL.XLSB in XLSTART and open it.
     * @returns {Promise<{ success: boolean, created: boolean, opened: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, fileExists: boolean, workbookPath?: string, message?: string }>}
     */
    create: () => ipcRenderer.invoke('personal:create')
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
  // MULTI-INSTANCE RESOLUTION
  // ==========================================================================
  // These APIs are used by renderer recovery flows when workbook polling detects
  // multi-instance/no-workbook states. Payloads intentionally include diagnostics
  // fields (pid/workbookCount/strategy/attempt) for log traceability.
  /**
   * Attempt to silently resolve the correct Excel instance
   * @returns {Promise<{
   *   resolved: boolean,
   *   reason?: string,
   *   pid?: number,
   *   workbookCount?: number,
   *   strategy?: 'foreground' | 'max_workbooks',
   *   attempt?: number,
   *   message?: string
   * }>}
   */
  resolveInstance: () => ipcRenderer.invoke('excel:resolveInstance'),

  /**
   * Clear COM cache and retry connection (for reconnect after user focuses correct Excel)
   * @returns {Promise<{ success: boolean, name?: string, path?: string }>}
   */
  reconnect: () => ipcRenderer.invoke('excel:reconnect'),

  // ==========================================================================
  // EVENT SUBSCRIPTIONS
  // ==========================================================================
  events: {
    /**
     * Subscribe to foreground Excel context transitions reported by WindowHelper.
     * @param {(payload: { excelActive: boolean, excelHwnd: string | null, process: string, timestamp: number }) => void} callback
     * @returns {() => void} unsubscribe function
     */
    onForegroundChanged: (callback) => {
      if (typeof callback !== 'function') {
        return () => {};
      }

      const listener = (_event, payload) => {
        callback(payload);
      };
      ipcRenderer.on('excel:foreground-changed', listener);
      return () => {
        ipcRenderer.removeListener('excel:foreground-changed', listener);
      };
    }
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
