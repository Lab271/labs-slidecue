import { autoUpdater } from 'electron-updater';
import { app, BrowserWindow, dialog } from 'electron';

// Configure for public releases repo
// Create a PUBLIC repo called "SlideCue-releases" to host your releases
// This allows auto-updates to work without exposing your private source code
const RELEASES_REPO = {
  owner: 'LAB271',
  repo: 'SlideCue-releases',
};

export function setupAutoUpdater(mainWindow: BrowserWindow) {
  // Skip updates in development
  if (!app.isPackaged) {
    console.log('Skip checkForUpdates because application is not packed and dev update config is not forced');
    return;
  }

  // Configure the update feed URL for the public releases repo
  autoUpdater.setFeedURL({
    provider: 'github',
    owner: RELEASES_REPO.owner,
    repo: RELEASES_REPO.repo,
  });

  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('checking-for-update', () => {
    console.log('Checking for updates...');
  });

  autoUpdater.on('update-available', async (info) => {
    console.log('Update available:', info.version);
    
    const result = await dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Available',
      message: `A new version (${info.version}) is available!`,
      detail: 'Would you like to download and install it now?',
      buttons: ['Download', 'Later'],
      defaultId: 0,
      cancelId: 1,
    });

    if (result.response === 0) {
      autoUpdater.downloadUpdate();
    }
  });

  autoUpdater.on('update-not-available', () => {
    console.log('No updates available - you have the latest version');
  });

  autoUpdater.on('download-progress', (progress) => {
    console.log(`Download progress: ${Math.round(progress.percent)}%`);
    mainWindow.setProgressBar(progress.percent / 100);
  });

  autoUpdater.on('update-downloaded', async (info) => {
    console.log('Update downloaded:', info.version);
    mainWindow.setProgressBar(-1); // Remove progress bar
    
    const result = await dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Update Ready',
      message: 'Update downloaded!',
      detail: 'The update will be installed when you restart SlideCue.',
      buttons: ['Restart Now', 'Later'],
      defaultId: 0,
      cancelId: 1,
    });

    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });

  autoUpdater.on('error', (error) => {
    console.error('Auto-updater error:', error.message);
    // Don't show error dialogs - just log them
  });

  // Check for updates after a short delay
  setTimeout(() => {
    autoUpdater.checkForUpdates().catch((err) => {
      console.error('Failed to check for updates:', err.message);
    });
  }, 3000);
}
