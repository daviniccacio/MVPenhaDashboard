const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    openExcel: () => ipcRenderer.invoke('open-excel'),
    getHistorico: () => ipcRenderer.invoke("get-historico"),
    openExcelPath: (filePath) => ipcRenderer.invoke("open-excel-path", filePath),
    setTitle: (title) => ipcRenderer.send("set-title", title),
});