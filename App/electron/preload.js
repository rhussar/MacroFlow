const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  testExcelConnection: () => ipcRenderer.invoke('test-excel-connection')
});
