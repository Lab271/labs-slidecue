# Architecture

## Overview

SlideCue is an Electron application that controls native Microsoft PowerPoint presentations while providing a web-based remote control interface accessible from any device on the local network.

```
┌─────────────────────────────────────────────────────────────────┐
│                       SLIDECUE APP                               │
│  ┌─────────────────┐    ┌─────────────────┐                     │
│  │  Main Process   │    │ Renderer Process│                     │
│  │                 │    │   (Home Screen) │                     │
│  │  - PowerPoint   │◄──►│   - Import UI   │                     │
│  │    Automation   │    │   - Thumbnails  │                     │
│  │  - Express      │    │   - PIN Display │                     │
│  │    Server       │    └─────────────────┘                     │
│  │  - Socket.IO    │    └─────────────────┘                     │
│  │  - Auto-Update  │                                            │
│  └────────┬────────┘                                            │
│           │                                                      │
└───────────┼──────────────────────────────────────────────────────┘
            │
            │ COM/AppleScript
            ▼
┌─────────────────────┐         ┌─────────────────────────────────┐
│  Microsoft          │         │     Web Remote (Phone/Tablet)   │
│  PowerPoint         │         │  ┌───────────────────────────┐  │
│  ┌───────────────┐  │         │  │   PIN Entry Screen        │  │
│  │  Slideshow    │  │         │  └───────────────────────────┘  │
│  │  (Primary     │  │         │              │                  │
│  │   Display)    │  │         │              ▼                  │
│  │               │  │  HTTP   │  ┌───────────────────────────┐  │
│  │  - Animations │◄─┼─────────┼──│   Remote Control UI       │  │
│  │  - Transitions│  │  WS     │  │   - Prev/Next Buttons     │  │
│  │  - Full PPTX  │  │         │  │   - Slide Preview         │  │
│  └───────────────┘  │         │  │   - Slide Counter         │  │
└─────────────────────┘         │  └───────────────────────────┘  │
                                └─────────────────────────────────┘
```

## Project Structure

```
slidecue/
├── .github/
│   └── workflows/
│       └── release.yml          # CI/CD for Windows builds
├── forge.config.ts              # Electron Forge configuration
├── electron.vite.config.ts      # Vite bundler configuration
├── package.json
├── tsconfig.json
├── tailwind.config.js
├── src/
│   ├── main/                    # Electron Main Process
│   │   ├── index.ts             # Entry point, window management
│   │   ├── updater.ts           # Auto-update logic
│   │   ├── pptx/
│   │   │   ├── automation.ts    # PowerPoint control abstraction
│   │   │   ├── macos.ts         # AppleScript implementation
│   │   │   └── windows.ts       # COM/winax implementation
│   │   └── server/
│   │       ├── index.ts         # Express server setup
│   │       └── socket.ts        # Socket.IO event handlers
│   ├── preload/
│   │   └── index.ts             # Context bridge (IPC)
│   ├── renderer/                # Electron Renderer (Desktop UI)
│   │   ├── index.html
│   │   └── src/
│   │       ├── main.tsx         # React entry
│   │       ├── App.tsx          # Main app component
│   │       ├── components/
│   │       │   ├── ImportDialog.tsx
│   │       │   ├── ThumbnailGrid.tsx
│   │       │   ├── ConnectionPanel.tsx
│   │       │   └── PowerPointCheck.tsx
│   │       ├── hooks/
│   │       │   └── usePresentation.ts
│   │       └── styles/
│   │           └── index.css    # Tailwind imports
│   └── remote/                  # Web Remote (Mobile UI)
│       ├── index.html
│       └── src/
│           ├── main.tsx
│           ├── App.tsx
│           ├── components/
│           │   ├── PinEntry.tsx
│           │   ├── SlideControls.tsx
│           │   ├── SlidePreview.tsx
│           │   └── SlideCounter.tsx
│           └── hooks/
│               └── useSocket.ts
└── resources/
    ├── icon.ico                 # Windows icon
    ├── icon.icns                # macOS icon
    └── icon.png                 # Linux icon
```

## Component Details

### Main Process (`src/main/`)

#### `index.ts` — Application Entry Point
- Creates the main `BrowserWindow` (home screen)
- Registers IPC handlers for renderer communication
- Manages application lifecycle (ready, quit, etc.)
- Checks PowerPoint installation on startup

#### `updater.ts` — Auto-Update System
- Uses `electron-updater` package
- Checks GitHub Releases on app startup
- Shows notification when update available
- Downloads and installs on user confirmation

#### `pptx/automation.ts` — PowerPoint Control Abstraction
```typescript
interface PowerPointAutomation {
  checkInstalled(): Promise<boolean>;
  openPresentation(filePath: string): Promise<void>;
  exportThumbnails(outputDir: string): Promise<string[]>;
  startSlideshow(): Promise<void>;
  nextSlide(): Promise<void>;
  prevSlide(): Promise<void>;
  gotoSlide(index: number): Promise<void>;
  getCurrentSlide(): Promise<number>;
  getTotalSlides(): Promise<number>;
  stopSlideshow(): Promise<void>;
}
```

#### `pptx/macos.ts` — AppleScript Implementation
Uses `child_process.exec` to run AppleScript commands:
```applescript
tell application "Microsoft PowerPoint"
    open POSIX file "/path/to/file.pptx"
    set theSettings to slide show settings of active presentation
    run slide show theSettings
end tell
```

#### `pptx/windows.ts` — COM/winax Implementation
Uses `winax` package for Windows COM automation:
```typescript
const winax = require('winax');
const ppt = new ActiveXObject('PowerPoint.Application');
ppt.Presentations.Open(filePath);
ppt.ActivePresentation.SlideShowSettings.Run();
```

#### `server/index.ts` — Express Server
- Serves the remote web UI on port 3000
- Serves slide thumbnails as static files
- CORS configured for local network access
- Binds to `0.0.0.0` for network accessibility

#### `server/socket.ts` — Socket.IO Handlers
- PIN validation on connection handshake
- Event handlers: `next`, `prev`, `goto`
- Broadcasts: `slide-changed`, `presentation-ended`
- Manages connected clients

### Preload Script (`src/preload/`)

#### `index.ts` — IPC Bridge
Exposes safe APIs to renderer via `contextBridge`:
```typescript
contextBridge.exposeInMainWorld('electronAPI', {
  importPresentation: () => ipcRenderer.invoke('import-presentation'),
  startPresentation: () => ipcRenderer.invoke('start-presentation'),
  stopPresentation: () => ipcRenderer.invoke('stop-presentation'),
  onSlideChanged: (callback) => ipcRenderer.on('slide-changed', callback),
  getConnectionInfo: () => ipcRenderer.invoke('get-connection-info'),
  isPowerPointInstalled: () => ipcRenderer.invoke('check-powerpoint'),
});
```

### Renderer — Desktop UI (`src/renderer/`)

#### `App.tsx` — Main Component
State machine with views:
1. **PowerPoint Check** — Shows error if not installed
2. **Import View** — File picker, no presentation loaded
3. **Preview View** — Thumbnails displayed, ready to present
4. **Presenting View** — Shows connection info (URL, PIN)

#### Key Components
- `PowerPointCheck` — Startup validation UI
- `ImportDialog` — File picker trigger
- `ThumbnailGrid` — Slide preview grid
- `ConnectionPanel` — URL, PIN display

### Remote — Web UI (`src/remote/`)

#### `App.tsx` — Remote Main Component
State machine:
1. **PIN Entry** — 4-digit PIN form
2. **Remote Control** — Slide navigation UI

#### Key Components
- `PinEntry` — PIN input form with validation
- `SlideControls` — Large Prev/Next buttons
- `SlidePreview` — Current or next slide thumbnail
- `SlideCounter` — "5 / 20" display

## Data Flow

### Presentation Start Flow
```
User clicks "Present"
        │
        ▼
Renderer ──IPC──► Main Process
                      │
                      ├──► Generate 4-digit PIN
                      ├──► Start Express server (port 3000)
                      ├──► Export slide thumbnails
                      ├──► Start PowerPoint slideshow
                      │
                      ▼
              Return connection info
                      │
        ◄─────────────┘
        │
        ▼
Display URL + PIN
```

### Slide Navigation Flow
```
Remote taps "Next"
        │
        ▼
Socket.IO emit('next')
        │
        ▼
Server validates PIN ──► Reject if invalid
        │
        ▼
Call automation.nextSlide()
        │
        ▼
PowerPoint advances
        │
        ▼
Server broadcasts('slide-changed', { current: 6, total: 20 })
        │
        ▼
All remotes update UI
```

## Security Model

### PIN Authentication
- Random 4-digit PIN generated per presentation session
- PIN validated on Socket.IO handshake
- Invalid PIN = connection rejected
- PIN displayed only on presenter's screen

### Network Scope
- Server binds to local network only
- No internet exposure (unless user configures router)
- mDNS/Bonjour for local discovery (optional)

## Build & Distribution

### Development
```bash
npm run dev          # Start with hot reload
```

### Production Build
```bash
npm run build        # Compile TypeScript
npm run make         # Create Windows installer
```

### CI/CD (GitHub Actions)
- Triggered on tag push (`v*.*.*`)
- Builds Windows x64 .exe
- Uploads to GitHub Releases
- Auto-updater pulls from Releases

## Dependencies

### Production
| Package | Purpose |
|---------|---------|
| `electron` | Desktop app framework |
| `express` | Web server |
| `socket.io` | Real-time communication |
| `socket.io-client` | Remote client communication |
| `winax` | Windows COM automation |
| `electron-updater` | Auto-updates |
| `bonjour-service` | Network discovery |

### Development
| Package | Purpose |
|---------|---------|
| `electron-vite` | Build tooling |
| `@electron-forge/*` | Packaging |
| `typescript` | Type safety |
| `tailwindcss` | Styling |
| `react` | UI framework |
