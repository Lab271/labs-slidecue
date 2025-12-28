import { contextBridge, ipcRenderer } from 'electron';

const electronAPI = {
  checkPowerPoint: () => ipcRenderer.invoke('check-powerpoint'),
  importPresentation: () => ipcRenderer.invoke('import-presentation'),
  startPresentation: () => ipcRenderer.invoke('start-presentation'),
  stopPresentation: () => ipcRenderer.invoke('stop-presentation'),
  getSlideInfo: () => ipcRenderer.invoke('get-slide-info'),
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);

// TypeScript declaration
export type ElectronAPI = typeof electronAPI;
