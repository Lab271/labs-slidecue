import { contextBridge, ipcRenderer } from 'electron';

const electronAPI = {
  checkPowerPoint: () => ipcRenderer.invoke('check-powerpoint'),
  importPresentation: () => ipcRenderer.invoke('import-presentation'),
  startPresentation: () => ipcRenderer.invoke('start-presentation'),
  stopPresentation: () => ipcRenderer.invoke('stop-presentation'),
  getSlideInfo: () => ipcRenderer.invoke('get-slide-info'),
  onImportProgress: (callback: (data: { step: number; total: number; message: string }) => void) => {
    ipcRenderer.on('import-progress', (_event, data) => callback(data));
    return () => ipcRenderer.removeAllListeners('import-progress');
  },
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);

// TypeScript declaration
export type ElectronAPI = typeof electronAPI;
