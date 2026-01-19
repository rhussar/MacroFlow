/**
 * Preload Script - Secure bridge between React UI and Electron main process
 *
 * This exposes a clean API to the renderer process via window.excel
 * All methods return Promises with { success: boolean, ... } responses
 *
 * Usage in React:
 *   const result = await window.excel.vba.inject({ code: '...' });
 *   const cell = await window.excel.cell.read({ address: 'A1' });
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

// Legacy API for backward compatibility (can be removed later)
contextBridge.exposeInMainWorld('electronAPI', {
  closeApp: () => ipcRenderer.send('app:close'),
  injectCode: (code) => ipcRenderer.invoke('vba:inject', { code }),
});
