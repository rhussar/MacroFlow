import { contextBridge, ipcRenderer } from 'electron'

/**
 * Preload script: Secure bridge between renderer and main process
 * Exposes only specific APIs to the renderer
 */

contextBridge.exposeInMainWorld('electronAPI', {
  testExcelConnection: () => ipcRenderer.invoke('test-excel-connection')
})

// Type definition for the exposed API
export interface ElectronAPI {
  testExcelConnection: () => Promise<{
    success: boolean
    message: string
  }>
}

declare global {
  interface Window {
    electronAPI: ElectronAPI
  }
}
