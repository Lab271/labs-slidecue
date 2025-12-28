import { autoUpdater } from 'electron-updater';
import { BrowserWindow, dialog } from 'electron';

export function setupAutoUpdater(mainWindow: BrowserWindow) {
  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('update-available', async (info) => {
    const result = await dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Available',
      message: `Version ${info.version} is available. Download now?`,
      buttons: ['Download', 'Later'],
      defaultId: 0,
    });

    if (result.response === 0) {
      autoUpdater.downloadUpdate();
    }
  });

  autoUpdater.on('update-downloaded', async () => {
    const result = await dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Ready',
      message: 'Update downloaded. Restart to install?',
      buttons: ['Restart', 'Later'],
      defaultId: 0,
    });

    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });

  autoUpdater.on('error', (error) => {
    console.error('Auto-update error:', error);
  });

  // Check for updates on startup
  autoUpdater.checkForUpdates();
}
