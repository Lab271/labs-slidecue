import { autoUpdater } from 'electron-updater';
import { app, BrowserWindow, dialog } from 'electron';
import log from 'electron-log';

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
    log.info('[Updater] Skip checkForUpdates because application is not packed');
    return;
  }

  // Configure the update feed URL for the public releases repo
  autoUpdater.setFeedURL({
    provider: 'github',
    owner: RELEASES_REPO.owner,
    repo: RELEASES_REPO.repo,
  });

  log.info('[Updater] Configured feed URL:', {
    provider: 'github',
    owner: RELEASES_REPO.owner,
    repo: RELEASES_REPO.repo,
  });

  autoUpdater.autoDownload = false;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('checking-for-update', () => {
    log.info('[Updater] Checking for updates...');
  });

  autoUpdater.on('update-available', async (info) => {
    log.info('[Updater] Update available:', JSON.stringify(info, null, 2));
    
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
      log.info('[Updater] User chose to download update');
      log.info('[Updater] Download URL:', info.files?.[0]?.url || 'No URL in info');
      autoUpdater.downloadUpdate().catch((err) => {
        log.error('[Updater] Failed to download update:', err);
        dialog.showMessageBox(mainWindow, {
          type: 'error',
          title: 'Download Failed',
          message: 'Failed to download update',
          detail: err.message,
        });
      });
    } else {
      log.info('[Updater] User chose to skip update');
    }
  });

  autoUpdater.on('update-not-available', () => {
    log.info('[Updater] No updates available - you have the latest version');
  });

  autoUpdater.on('download-progress', (progress) => {
    log.info(`[Updater] Download progress: ${Math.round(progress.percent)}%`);
    mainWindow.setProgressBar(progress.percent / 100);
  });

  autoUpdater.on('update-downloaded', async (info) => {
    log.info('[Updater] Update downloaded:', info.version);
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
      log.info('[Updater] User chose to restart now');
      autoUpdater.quitAndInstall();
    } else {
      log.info('[Updater] User chose to restart later');
    }
  });

  autoUpdater.on('error', (error) => {
    log.error('[Updater] Auto-updater error:', error.message, error.stack);
    // Only log errors, don't show dialog to user as it's disruptive
    // Network errors and 404s are expected when releases repo doesn't have new versions
  });

  // Check for updates after a short delay
  setTimeout(() => {
    log.info('[Updater] Initiating update check');
    autoUpdater.checkForUpdates().catch((err) => {
      log.error('[Updater] Failed to check for updates:', err.message);
    });
  }, 3000);
}
