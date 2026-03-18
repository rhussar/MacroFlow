const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('installer', {
  openExcel: () => ipcRenderer.send('first-run:open-excel'),
  done: () => ipcRenderer.send('first-run:done'),
});
