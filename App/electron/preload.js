const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  testExcelConnection: () => ipcRenderer.invoke('test-excel-connection'),
  closeApp: () => ipcRenderer.send('close-app'),
  injectCode: (code) => ipcRenderer.invoke('vba:inject', code)
});
