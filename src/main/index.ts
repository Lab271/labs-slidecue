// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { app, BrowserWindow, ipcMain, dialog, protocol, net } from 'electron';
import log from 'electron-log';
import path from 'path';
import os from 'os';
import fs from 'fs/promises';
import fsSync from 'fs';
import { getAutomation } from './pptx/automation';
import { startServer, stopServer, getLocalIP } from './server';
import { setupSocketHandlers } from './server/socket';
import { setupAutoUpdater } from './updater';

// Disable hardware acceleration on Windows to prevent mouse lag
if (process.platform === 'win32') {
  app.disableHardwareAcceleration();
}

let mainWindow: BrowserWindow | null = null;
let presentationFile: string | null = null;
let thumbnailPaths: string[] = [];
let currentThumbsDir: string = '';

const automation = getAutomation();

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    fullscreen: false,
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: false, // Allow loading local files
    },
  });

  if (process.env.NODE_ENV === 'development') {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
  }

  return mainWindow;
}

// Register custom protocol to serve local files
protocol.registerSchemesAsPrivileged([
  { scheme: 'slidecue', privileges: { secure: true, supportFetchAPI: true, stream: true } }
]);

app.whenReady().then(async () => {
  // Register protocol handler for local file access
  protocol.handle('slidecue', (request) => {
    const filePath = request.url.replace('slidecue://', '');
    return net.fetch('file://' + filePath);
  });

  const window = createWindow();

  // Check PowerPoint on startup (non-blocking)
  automation.checkInstalled().then(isPPTInstalled => {
    if (!isPPTInstalled) {
      dialog.showErrorBox(
        'PowerPoint Not Found',
        'Microsoft PowerPoint is required to run presentations. Please install PowerPoint and restart SlideCue.'
      );
    }
  });

  // Setup auto-updater (production only)
  if (process.env.NODE_ENV !== 'development') {
    setupAutoUpdater(window);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// IPC Handlers
ipcMain.handle('check-powerpoint', async () => {
  return automation.checkInstalled();
});

ipcMain.handle('import-presentation', async () => {
  const result = await dialog.showOpenDialog(mainWindow!, {
    filters: [{ name: 'PowerPoint', extensions: ['pptx', 'ppt'] }],
    properties: ['openFile'],
  });

  if (result.canceled || !result.filePaths[0]) {
    return null;
  }

  presentationFile = result.filePaths[0];
  log.info('Selected presentation:', presentationFile);

  // Create directory for thumbnails in app userData
  const appDataDir = path.join(app.getPath('userData'), 'thumbnails');
  currentThumbsDir = path.join(
    appDataDir,
    Date.now().toString()
  );
  log.info('Creating thumbnails directory:', currentThumbsDir);
  await fs.mkdir(currentThumbsDir, { recursive: true });

  // Send progress callback
  const sendProgress = (step: number, total: number, message: string) => {
    mainWindow?.webContents.send('import-progress', { step, total, message });
  };

  // Now that we have a file selected, start the loading screen
  sendProgress(0, 5, 'Starting...');
  log.info('Starting import process');
  
  // Export thumbnails with progress
  sendProgress(1, 5, 'Opening presentation...');
  log.info('Opening presentation in PowerPoint');
  await automation.openPresentation(presentationFile);
  
  sendProgress(2, 5, 'Getting slide count...');
  await automation.getSlideInfo();
  
  sendProgress(3, 5, 'Converting to PDF...');
  const slideMetadata = await automation.exportThumbnails(currentThumbsDir, (current, total) => {
    sendProgress(3 + (current / total), 5, `Converting slide ${current}/${total}...`);
  });
  
  thumbnailPaths = slideMetadata.thumbnails;
  sendProgress(5, 5, 'Done!');

  console.log('Thumbnails exported to:', currentThumbsDir);
  console.log('Thumbnail files:', thumbnailPaths);
  console.log('Hidden slides:', slideMetadata.hiddenSlides);

  return {
    filePath: presentationFile,
    fileName: path.basename(presentationFile),
    thumbnails: thumbnailPaths,
    totalSlides: slideMetadata.totalSlides,
    visibleSlides: slideMetadata.visibleSlides,
    hiddenSlides: slideMetadata.hiddenSlides,
  };
});

ipcMain.handle('start-presentation', async () => {
  if (!presentationFile) {
    throw new Error('No presentation loaded');
  }

  // Start web server - use resources folder for remote UI
  const remoteUIPath = app.isPackaged
    ? path.join(process.resourcesPath, 'remote')
    : path.join(app.getAppPath(), 'resources/remote');
  
  console.log('Remote UI path:', remoteUIPath);
  console.log('Thumbnails dir:', currentThumbsDir);
  
  const { url, pin, io } = await startServer(remoteUIPath, currentThumbsDir);

  // Setup socket handlers
  setupSocketHandlers(io);

  return {
    url,
    pin,
    localIP: getLocalIP(),
  };
});

ipcMain.handle('stop-presentation', async () => {
  await automation.stopSlideshow();
  stopServer();
});

ipcMain.handle('start-slideshow', async () => {
  await automation.startSlideshow();
});

ipcMain.handle('get-slide-info', async () => {
  return automation.getSlideInfo();
});

// Cleanup temp directories (synchronous for use in quit handler)
function cleanupTempDirs() {
  log.info('Cleaning up thumbnail directories');
  const thumbsDir = path.join(app.getPath('userData'), 'thumbnails');
  const tempPresentationsDir = path.join(os.tmpdir(), 'slidecue-presentations');
  
  // Clean all thumbnail subdirectories
  try {
    if (fsSync.existsSync(thumbsDir)) {
      const entries = fsSync.readdirSync(thumbsDir);
      for (const entry of entries) {
        const entryPath = path.join(thumbsDir, entry);
        fsSync.rmSync(entryPath, { recursive: true, force: true });
        log.info(`Cleaned up: ${entryPath}`);
      }
    }
  } catch (e) {
    log.error('Error cleaning thumbnails:', e);
  }
  
  // Clean temp presentations
  try {
    if (fsSync.existsSync(tempPresentationsDir)) {
      fsSync.rmSync(tempPresentationsDir, { recursive: true, force: true });
      log.info(`Cleaned up: ${tempPresentationsDir}`);
    }
  } catch (e) {
    log.error('Error cleaning temp presentations:', e);
  }
}

app.on('will-quit', () => {
  cleanupTempDirs();
});

app.on('window-all-closed', () => {
  stopServer();
  // Quit on all platforms (override macOS default behavior)
  app.quit();
});
