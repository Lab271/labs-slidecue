import { app, BrowserWindow, ipcMain, dialog, protocol, net } from 'electron';
import path from 'path';
import os from 'os';
import fs from 'fs/promises';
import { getAutomation } from './pptx/automation';
import { startServer, stopServer, getLocalIP } from './server';
import { setupSocketHandlers } from './server/socket';
import { setupAutoUpdater } from './updater';

let mainWindow: BrowserWindow | null = null;
let presentationFile: string | null = null;
let thumbnailPaths: string[] = [];
let currentThumbsDir: string = '';

const automation = getAutomation();

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
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

  // Check PowerPoint on startup
  const isPPTInstalled = await automation.checkInstalled();
  if (!isPPTInstalled) {
    dialog.showErrorBox(
      'PowerPoint Not Found',
      'Microsoft PowerPoint is required to run presentations. Please install PowerPoint and restart SlideCue.'
    );
  }

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

  // Create temp directory for thumbnails
  currentThumbsDir = path.join(
    os.tmpdir(),
    'slidecue-thumbs',
    Date.now().toString()
  );
  await fs.mkdir(currentThumbsDir, { recursive: true });

  // Export thumbnails
  await automation.openPresentation(presentationFile);
  thumbnailPaths = await automation.exportThumbnails(currentThumbsDir);

  console.log('Thumbnails exported to:', currentThumbsDir);
  console.log('Thumbnail files:', thumbnailPaths);

  return {
    filePath: presentationFile,
    fileName: path.basename(presentationFile),
    thumbnails: thumbnailPaths,
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

  // Start PowerPoint slideshow
  await automation.startSlideshow();

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

ipcMain.handle('get-slide-info', async () => {
  return automation.getSlideInfo();
});

app.on('window-all-closed', () => {
  stopServer();
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
