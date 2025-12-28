# Implementation Guide

Step-by-step instructions to build SlideCue from scratch.

## Phase 1: Project Setup

### Step 1.1: Initialize Electron Project

```bash
# Create project with electron-vite and React TypeScript template
npm create @aspect-build/electron-vite@latest slidecue -- --template=react-ts
cd slidecue

# Or use Electron Forge with Vite
npm init electron-app@latest slidecue -- --template=vite-typescript
cd slidecue
```

### Step 1.2: Install Dependencies

```bash
# Production dependencies
npm install express socket.io socket.io-client bonjour-service electron-updater

# Windows COM automation (optional on macOS, required on Windows)
npm install winax --save-optional

# Dev dependencies
npm install -D tailwindcss postcss autoprefixer @types/express
```

### Step 1.3: Configure Tailwind CSS

```bash
npx tailwindcss init -p
```

Update `tailwind.config.js`:
```javascript
/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './src/renderer/**/*.{html,tsx,ts}',
    './src/remote/**/*.{html,tsx,ts}',
  ],
  theme: {
    extend: {},
  },
  plugins: [],
}
```

Create `src/renderer/src/styles/index.css`:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;
```

### Step 1.4: Configure Electron Forge for Windows

Update `forge.config.ts`:
```typescript
import type { ForgeConfig } from '@electron-forge/shared-types';
import { MakerSquirrel } from '@electron-forge/maker-squirrel';

const config: ForgeConfig = {
  packagerConfig: {
    icon: './resources/icon',
    appBundleId: 'com.yourname.slidecue',
  },
  makers: [
    new MakerSquirrel({
      setupIcon: './resources/icon.ico',
      iconUrl: 'https://raw.githubusercontent.com/yourusername/slidecue/main/resources/icon.ico',
    }),
  ],
  publishers: [
    {
      name: '@electron-forge/publisher-github',
      config: {
        repository: {
          owner: 'yourusername',
          name: 'slidecue',
        },
        prerelease: false,
      },
    },
  ],
};

export default config;
```

---

## Phase 2: PowerPoint Automation

### Step 2.1: Create Automation Interface

Create `src/main/pptx/types.ts`:
```typescript
export interface SlideInfo {
  current: number;
  total: number;
}

export interface PowerPointAutomation {
  checkInstalled(): Promise<boolean>;
  openPresentation(filePath: string): Promise<void>;
  exportThumbnails(outputDir: string): Promise<string[]>;
  startSlideshow(): Promise<void>;
  nextSlide(): Promise<void>;
  prevSlide(): Promise<void>;
  gotoSlide(index: number): Promise<void>;
  getSlideInfo(): Promise<SlideInfo>;
  stopSlideshow(): Promise<void>;
  closePresentation(): Promise<void>;
}
```

### Step 2.2: macOS Implementation (AppleScript)

Create `src/main/pptx/macos.ts`:
```typescript
import { exec } from 'child_process';
import { promisify } from 'util';
import { PowerPointAutomation, SlideInfo } from './types';

const execAsync = promisify(exec);

function runAppleScript(script: string): Promise<string> {
  return execAsync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`)
    .then(({ stdout }) => stdout.trim());
}

export const macOSAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      const result = await runAppleScript('tell application "System Events" to exists application file "Microsoft PowerPoint.app" of folder "Applications" of startup disk');
      return result === 'true';
    } catch {
      return false;
    }
  },

  async openPresentation(filePath: string) {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        activate
        open POSIX file "${filePath}"
      end tell
    `);
  },

  async exportThumbnails(outputDir: string) {
    // PowerPoint exports slides as images into a folder
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        save active presentation in POSIX file "${outputDir}" as save as PNG
      end tell
    `);
    // PowerPoint creates a folder with the presentation name containing slide images
    const fs = await import('fs/promises');
    const path = await import('path');
    const files = await fs.readdir(outputDir);
    const pngFiles = files
      .filter(f => f.endsWith('.png'))
      .sort()
      .map(f => path.join(outputDir, f));
    return pngFiles;
  },

  async startSlideshow() {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        activate
        set theSettings to slide show settings of active presentation
        run slide show theSettings
      end tell
    `);
  },

  async nextSlide() {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        tell slide show view of slide show window 1
          go to next slide
        end tell
      end tell
    `);
  },

  async prevSlide() {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        tell slide show view of slide show window 1
          go to previous slide
        end tell
      end tell
    `);
  },

  async gotoSlide(index: number) {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        tell slide show view of slide show window 1
          go to slide ${index}
        end tell
      end tell
    `);
  },

  async getSlideInfo(): Promise<SlideInfo> {
    const current = await runAppleScript(`
      tell application "Microsoft PowerPoint"
        tell slide show view of slide show window 1
          return slide index of current slide
        end tell
      end tell
    `);
    const total = await runAppleScript(`
      tell application "Microsoft PowerPoint"
        return count of slides of active presentation
      end tell
    `);
    return {
      current: parseInt(current, 10),
      total: parseInt(total, 10),
    };
  },

  async stopSlideshow() {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        end slide show of slide show window 1
      end tell
    `);
  },

  async closePresentation() {
    await runAppleScript(`
      tell application "Microsoft PowerPoint"
        close active presentation saving no
      end tell
    `);
  },
};
```

### Step 2.3: Windows Implementation (winax/COM)

Create `src/main/pptx/windows.ts`:
```typescript
import { PowerPointAutomation, SlideInfo } from './types';

// winax is Windows-only, conditionally import
let winax: any;
try {
  winax = require('winax');
} catch {
  // Not on Windows, will throw if methods are called
}

let pptApp: any = null;
let presentation: any = null;
let slideShow: any = null;

export const windowsAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      const testApp = new winax.Object('PowerPoint.Application');
      testApp.Quit();
      return true;
    } catch {
      return false;
    }
  },

  async openPresentation(filePath: string) {
    pptApp = new winax.Object('PowerPoint.Application');
    pptApp.Visible = true;
    presentation = pptApp.Presentations.Open(filePath);
  },

  async exportThumbnails(outputDir: string) {
    const paths: string[] = [];
    const count = presentation.Slides.Count;
    
    for (let i = 1; i <= count; i++) {
      const slide = presentation.Slides.Item(i);
      const filePath = `${outputDir}/slide_${i.toString().padStart(3, '0')}.png`;
      slide.Export(filePath, 'PNG', 1920, 1080);
      paths.push(filePath);
    }
    
    return paths;
  },

  async startSlideshow() {
    const settings = presentation.SlideShowSettings;
    settings.StartingSlide = 1;
    settings.EndingSlide = presentation.Slides.Count;
    slideShow = settings.Run();
  },

  async nextSlide() {
    if (slideShow?.View) {
      slideShow.View.Next();
    }
  },

  async prevSlide() {
    if (slideShow?.View) {
      slideShow.View.Previous();
    }
  },

  async gotoSlide(index: number) {
    if (slideShow?.View) {
      slideShow.View.GotoSlide(index);
    }
  },

  async getSlideInfo(): Promise<SlideInfo> {
    return {
      current: slideShow?.View?.CurrentShowPosition || 1,
      total: presentation?.Slides?.Count || 0,
    };
  },

  async stopSlideshow() {
    if (slideShow?.View) {
      slideShow.View.Exit();
      slideShow = null;
    }
  },

  async closePresentation() {
    if (presentation) {
      presentation.Close();
      presentation = null;
    }
    if (pptApp) {
      pptApp.Quit();
      pptApp = null;
    }
  },
};
```

### Step 2.4: Platform Abstraction

Create `src/main/pptx/automation.ts`:
```typescript
import { PowerPointAutomation } from './types';
import { macOSAutomation } from './macos';
import { windowsAutomation } from './windows';

export function getAutomation(): PowerPointAutomation {
  if (process.platform === 'darwin') {
    return macOSAutomation;
  } else if (process.platform === 'win32') {
    return windowsAutomation;
  }
  throw new Error('Unsupported platform');
}

export * from './types';
```

---

## Phase 3: Web Server & Socket.IO

### Step 3.1: Express Server Setup

Create `src/main/server/index.ts`:
```typescript
import express from 'express';
import { createServer } from 'http';
import { Server } from 'socket.io';
import path from 'path';
import { networkInterfaces } from 'os';

let server: ReturnType<typeof createServer> | null = null;
let io: Server | null = null;
let currentPin: string = '';
let thumbnailsDir: string = '';
let currentPort: number = 3000;

function isPortAvailable(port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const testServer = createServer();
    testServer.once('error', () => resolve(false));
    testServer.once('listening', () => {
      testServer.close();
      resolve(true);
    });
    testServer.listen(port, '0.0.0.0');
  });
}

async function findAvailablePort(startPort: number): Promise<number> {
  let port = startPort;
  while (!(await isPortAvailable(port))) {
    port++;
  }
  return port;
}

export function getLocalIP(): string {
  const nets = networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name] || []) {
      if (net.family === 'IPv4' && !net.internal) {
        return net.address;
      }
    }
  }
  return 'localhost';
}

export function generatePin(): string {
  return Math.floor(1000 + Math.random() * 9000).toString();
}

export async function startServer(remoteUIPath: string, thumbsDir: string): Promise<{ 
  url: string; 
  pin: string;
  io: Server;
}> {
  const app = express();
  server = createServer(app);
  io = new Server(server, {
    cors: { origin: '*' },
  });

  currentPin = generatePin();
  thumbnailsDir = thumbsDir;
  currentPort = await findAvailablePort(3000);

  // Serve remote UI
  app.use('/remote', express.static(remoteUIPath));
  
  // Serve slide thumbnails
  app.use('/thumbnails', express.static(thumbnailsDir));

  // Socket.IO with PIN auth
  io.use((socket, next) => {
    const pin = socket.handshake.auth.pin;
    if (pin === currentPin) {
      next();
    } else {
      next(new Error('Invalid PIN'));
    }
  });

  server.listen(currentPort, '0.0.0.0');

  const localIP = getLocalIP();
  return {
    url: `http://${localIP}:${currentPort}/remote`,
    pin: currentPin,
    io,
  };
}

export function stopServer() {
  if (io) {
    io.close();
    io = null;
  }
  if (server) {
    server.close();
    server = null;
  }
}

export function broadcastSlideChange(current: number, total: number) {
  if (io) {
    io.emit('slide-changed', { current, total });
  }
}
```

### Step 3.2: Socket Event Handlers

Create `src/main/server/socket.ts`:
```typescript
import { Server } from 'socket.io';
import { getAutomation } from '../pptx/automation';

export function setupSocketHandlers(io: Server) {
  const automation = getAutomation();

  io.on('connection', (socket) => {
    console.log('Remote connected:', socket.id);

    // Send current slide info on connect
    automation.getSlideInfo().then((info) => {
      socket.emit('slide-changed', info);
    });

    socket.on('next', async () => {
      await automation.nextSlide();
      const info = await automation.getSlideInfo();
      io.emit('slide-changed', info);
    });

    socket.on('prev', async () => {
      await automation.prevSlide();
      const info = await automation.getSlideInfo();
      io.emit('slide-changed', info);
    });

    socket.on('goto', async (slideIndex: number) => {
      await automation.gotoSlide(slideIndex);
      const info = await automation.getSlideInfo();
      io.emit('slide-changed', info);
    });

    socket.on('disconnect', () => {
      console.log('Remote disconnected:', socket.id);
    });
  });
}
```

---

## Phase 4: Auto-Updater

### Step 4.1: Updater Module

Create `src/main/updater.ts`:
```typescript
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
```

---

## Phase 5: Main Process Entry

### Step 5.1: Main Entry Point

Create/update `src/main/index.ts`:
```typescript
import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'path';
import os from 'os';
import { getAutomation } from './pptx/automation';
import { startServer, stopServer, broadcastSlideChange, getLocalIP } from './server';
import { setupSocketHandlers } from './server/socket';
import { setupAutoUpdater } from './updater';

let mainWindow: BrowserWindow | null = null;
let presentationFile: string | null = null;
let thumbnailPaths: string[] = [];

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
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));
  }

  return mainWindow;
}

app.whenReady().then(async () => {
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
  
  // Export thumbnails
  const thumbsDir = path.join(os.tmpdir(), 'slidecue-thumbs', Date.now().toString());
  await automation.openPresentation(presentationFile);
  thumbnailPaths = await automation.exportThumbnails(thumbsDir);

  return {
    filePath: presentationFile,
    thumbnails: thumbnailPaths,
  };
});

ipcMain.handle('start-presentation', async () => {
  if (!presentationFile) {
    throw new Error('No presentation loaded');
  }

  // Start web server
  const remoteUIPath = path.join(__dirname, '../remote');
  const thumbsDir = path.dirname(thumbnailPaths[0] || '');
  const { url, pin, io } = await startServer(remoteUIPath, thumbsDir);

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
```

---

## Phase 6: Preload Script

### Step 6.1: IPC Bridge

Create/update `src/preload/index.ts`:
```typescript
import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
  checkPowerPoint: () => ipcRenderer.invoke('check-powerpoint'),
  importPresentation: () => ipcRenderer.invoke('import-presentation'),
  startPresentation: () => ipcRenderer.invoke('start-presentation'),
  stopPresentation: () => ipcRenderer.invoke('stop-presentation'),
  getSlideInfo: () => ipcRenderer.invoke('get-slide-info'),
});

// TypeScript declaration
declare global {
  interface Window {
    electronAPI: {
      checkPowerPoint: () => Promise<boolean>;
      importPresentation: () => Promise<{ filePath: string; thumbnails: string[] } | null>;
      startPresentation: () => Promise<{ url: string; pin: string; localIP: string }>;
      stopPresentation: () => Promise<void>;
      getSlideInfo: () => Promise<{ current: number; total: number }>;
    };
  }
}
```

---

## Phase 7: Renderer UI (Desktop)

### Step 7.1: Main App Component

Create `src/renderer/src/App.tsx`:
```tsx
import { useState, useEffect } from 'react';
import './styles/index.css';

type AppState = 'checking' | 'no-powerpoint' | 'idle' | 'loaded' | 'presenting';

interface PresentationInfo {
  filePath: string;
  thumbnails: string[];
}

interface ConnectionInfo {
  url: string;
  pin: string;
}

export default function App() {
  const [state, setState] = useState<AppState>('checking');
  const [presentation, setPresentation] = useState<PresentationInfo | null>(null);
  const [connection, setConnection] = useState<ConnectionInfo | null>(null);

  useEffect(() => {
    window.electronAPI.checkPowerPoint().then((installed) => {
      setState(installed ? 'idle' : 'no-powerpoint');
    });
  }, []);

  const handleImport = async () => {
    const result = await window.electronAPI.importPresentation();
    if (result) {
      setPresentation(result);
      setState('loaded');
    }
  };

  const handlePresent = async () => {
    const result = await window.electronAPI.startPresentation();
    setConnection(result);
    setState('presenting');
  };

  const handleStop = async () => {
    await window.electronAPI.stopPresentation();
    setConnection(null);
    setState('loaded');
  };

  if (state === 'checking') {
    return (
      <div className="flex items-center justify-center h-screen">
        <p className="text-xl">Checking PowerPoint...</p>
      </div>
    );
  }

  if (state === 'no-powerpoint') {
    return (
      <div className="flex flex-col items-center justify-center h-screen gap-4">
        <h1 className="text-2xl font-bold text-red-600">PowerPoint Not Found</h1>
        <p>Microsoft PowerPoint is required to use SlideCue.</p>
        <p>Please install PowerPoint and restart the app.</p>
      </div>
    );
  }

  if (state === 'presenting' && connection) {
    return (
      <div className="flex flex-col items-center justify-center h-screen gap-6 p-8">
        <h1 className="text-3xl font-bold">Presenting</h1>
        <p className="text-lg">Connect from any device on your network:</p>
        <p className="text-xl font-mono bg-gray-100 px-4 py-2 rounded">{connection.url}</p>
        <p className="text-6xl font-bold tracking-widest mt-4">{connection.pin}</p>
        <p className="text-gray-500">Enter this PIN on the remote</p>
        <button
          onClick={handleStop}
          className="mt-8 px-6 py-3 bg-red-500 text-white rounded-lg hover:bg-red-600"
        >
          Stop Presentation
        </button>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen p-8">
      <header className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">SlideCue</h1>
        <div className="flex gap-4">
          <button
            onClick={handleImport}
            className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
          >
            Import PPTX
          </button>
          {state === 'loaded' && (
            <button
              onClick={handlePresent}
              className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
            >
              Present
            </button>
          )}
        </div>
      </header>

      {presentation ? (
        <div className="grid grid-cols-4 gap-4 overflow-auto">
          {presentation.thumbnails.map((thumb, i) => (
            <div key={i} className="border rounded overflow-hidden">
              <img src={`file://${thumb}`} alt={`Slide ${i + 1}`} className="w-full" />
              <p className="text-center py-2 bg-gray-100">Slide {i + 1}</p>
            </div>
          ))}
        </div>
      ) : (
        <div className="flex-1 flex items-center justify-center border-2 border-dashed border-gray-300 rounded-lg">
          <p className="text-gray-500 text-xl">Import a PowerPoint file to get started</p>
        </div>
      )}
    </div>
  );
}
```

---

## Phase 8: Remote UI (Web)

### Step 8.1: Remote Entry HTML

Create `src/remote/index.html`:
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>SlideCue Remote</title>
</head>
<body>
  <div id="root"></div>
  <script type="module" src="/src/main.tsx"></script>
</body>
</html>
```

### Step 8.2: Remote App Component

Create `src/remote/src/App.tsx`:
```tsx
import { useState, useEffect } from 'react';
import { io, Socket } from 'socket.io-client';
import './styles/index.css';

type ViewMode = 'current' | 'next';

export default function App() {
  const [pin, setPin] = useState('');
  const [authenticated, setAuthenticated] = useState(false);
  const [error, setError] = useState('');
  const [socket, setSocket] = useState<Socket | null>(null);
  const [slideInfo, setSlideInfo] = useState({ current: 1, total: 1 });
  const [viewMode, setViewMode] = useState<ViewMode>('next');

  const handleConnect = () => {
    const newSocket = io(window.location.origin, {
      auth: { pin },
    });

    newSocket.on('connect', () => {
      setAuthenticated(true);
      setError('');
    });

    newSocket.on('connect_error', (err) => {
      setError('Invalid PIN');
      newSocket.close();
    });

    newSocket.on('slide-changed', (info) => {
      setSlideInfo(info);
    });

    setSocket(newSocket);
  };

  const handleNext = () => socket?.emit('next');
  const handlePrev = () => socket?.emit('prev');

  if (!authenticated) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen p-4 bg-gray-100">
        <h1 className="text-2xl font-bold mb-8">SlideCue Remote</h1>
        <input
          type="text"
          inputMode="numeric"
          pattern="[0-9]*"
          maxLength={4}
          placeholder="Enter 4-digit PIN"
          value={pin}
          onChange={(e) => setPin(e.target.value.replace(/\D/g, ''))}
          className="text-center text-3xl w-48 px-4 py-3 border-2 rounded-lg mb-4"
        />
        {error && <p className="text-red-500 mb-4">{error}</p>}
        <button
          onClick={handleConnect}
          disabled={pin.length !== 4}
          className="px-8 py-3 bg-blue-500 text-white rounded-lg text-xl disabled:opacity-50"
        >
          Connect
        </button>
      </div>
    );
  }

  const previewSlide = viewMode === 'current' ? slideInfo.current : Math.min(slideInfo.current + 1, slideInfo.total);

  return (
    <div className="flex flex-col min-h-screen bg-gray-100">
      {/* Slide Preview */}
      <div className="flex-1 p-4">
        <div className="bg-white rounded-lg shadow overflow-hidden">
          <img
            src={`/thumbnails/slide_${previewSlide.toString().padStart(3, '0')}.png`}
            alt={`Slide ${previewSlide}`}
            className="w-full"
          />
        </div>
        
        {/* View Toggle */}
        <div className="flex justify-center gap-4 mt-4">
          <button
            onClick={() => setViewMode('current')}
            className={`px-4 py-2 rounded ${viewMode === 'current' ? 'bg-blue-500 text-white' : 'bg-white'}`}
          >
            Current Slide
          </button>
          <button
            onClick={() => setViewMode('next')}
            className={`px-4 py-2 rounded ${viewMode === 'next' ? 'bg-blue-500 text-white' : 'bg-white'}`}
          >
            Next Slide
          </button>
        </div>
      </div>

      {/* Slide Counter */}
      <div className="text-center py-4">
        <span className="text-2xl font-bold">{slideInfo.current} / {slideInfo.total}</span>
      </div>

      {/* Navigation Buttons */}
      <div className="grid grid-cols-2 gap-4 p-4">
        <button
          onClick={handlePrev}
          disabled={slideInfo.current <= 1}
          className="py-8 bg-gray-800 text-white text-2xl rounded-lg disabled:opacity-50"
        >
          ← Previous
        </button>
        <button
          onClick={handleNext}
          disabled={slideInfo.current >= slideInfo.total}
          className="py-8 bg-blue-600 text-white text-2xl rounded-lg disabled:opacity-50"
        >
          Next →
        </button>
      </div>
    </div>
  );
}
```

---

## Phase 9: CI/CD Setup

### Step 9.1: GitHub Actions Workflow

Create `.github/workflows/release.yml`:
```yaml
name: Release

on:
  push:
    tags:
      - 'v*.*.*'

jobs:
  build-windows:
    runs-on: windows-latest
    
    steps:
      - uses: actions/checkout@v4
      
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: 20
          cache: 'npm'
      
      - name: Install dependencies
        run: npm ci
      
      - name: Build and Make
        run: npm run make
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
      
      - name: Upload artifact
        uses: actions/upload-artifact@v4
        with:
          name: windows-installer
          path: out/make/squirrel.windows/x64/*.exe

  release:
    needs: build-windows
    runs-on: ubuntu-latest
    permissions:
      contents: write
    
    steps:
      - name: Download artifacts
        uses: actions/download-artifact@v4
        with:
          name: windows-installer
          path: ./dist
      
      - name: Create Release
        uses: softprops/action-gh-release@v1
        with:
          files: ./dist/*.exe
          generate_release_notes: true
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

---

## Testing Checklist

- [ ] PowerPoint detection works on macOS and Windows
- [ ] PPTX import and thumbnail generation
- [ ] Slideshow starts and displays on primary monitor
- [ ] Web server accessible from other devices on network
- [ ] PIN validation rejects invalid codes
- [ ] Next/Previous commands control PowerPoint
- [ ] Slide changes broadcast to all remotes
- [ ] Current/Next slide preview toggles correctly
- [ ] Auto-updater detects new releases
- [ ] Windows installer (.exe) builds successfully
- [ ] App installs and runs on clean Windows machine
