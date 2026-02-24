const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    openExcel: () => ipcRenderer.invoke('open-excel')
});