"use strict";
const electron = require("electron");
const electronAPI = {
  checkPowerPoint: () => electron.ipcRenderer.invoke("check-powerpoint"),
  importPresentation: () => electron.ipcRenderer.invoke("import-presentation"),
  startPresentation: () => electron.ipcRenderer.invoke("start-presentation"),
  stopPresentation: () => electron.ipcRenderer.invoke("stop-presentation"),
  getSlideInfo: () => electron.ipcRenderer.invoke("get-slide-info"),
  onImportProgress: (callback) => {
    electron.ipcRenderer.on("import-progress", (_event, data) => callback(data));
    return () => electron.ipcRenderer.removeAllListeners("import-progress");
  }
};
electron.contextBridge.exposeInMainWorld("electronAPI", electronAPI);
