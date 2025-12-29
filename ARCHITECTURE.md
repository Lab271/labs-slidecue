# Architecture

## Overview

SlideCue is an Electron application that controls Microsoft PowerPoint presentations while providing a web-based remote control interface accessible from any device on the local network.

```
┌─────────────────────────────────────────────────────────────────┐
│                       SLIDECUE APP                               │
│  ┌─────────────────┐    ┌─────────────────┐                     │
│  │  Main Process   │    │ Renderer Process│                     │
│  │                 │    │   (Desktop UI)  │                     │
│  │  - PowerPoint   │◄──►│   - Import UI   │                     │
│  │    Automation   │IPC │   - Thumbnails  │                     │
│  │  - PPTX Parser  │    │   - QR Code     │                     │
│  │  - Express      │    │   - PIN Display │                     │
│  │    Server       │    └─────────────────┘                     │
│  │  - Socket.IO    │                                            │
│  └────────┬────────┘                                            │
│           │                                                      │
└───────────┼──────────────────────────────────────────────────────┘
            │
            │ AppleScript (macOS) / COM (Windows)
            ▼
┌─────────────────────┐         ┌─────────────────────────────────┐
│  Microsoft          │         │     Web Remote (Phone/Tablet)   │
│  PowerPoint         │         │  ┌───────────────────────────┐  │
│  ┌───────────────┐  │         │  │   PIN Entry / QR Scan     │  │
│  │  Slideshow    │  │         │  └───────────────────────────┘  │
│  │  Mode         │  │         │              │                  │
│  │               │  │  HTTP   │              ▼                  │
│  │  - Animations │◄─┼─────────┼──┌───────────────────────────┐  │
│  │  - Transitions│  │  WS     │  │   Remote Control UI       │  │
│  │  - Videos     │  │         │  │   - Current/Next Slides   │  │
│  └───────────────┘  │         │  │   - Speaker Notes         │  │
└─────────────────────┘         │  │   - Prev/Next/Goto        │  │
                                │  │   - Timer & Clock         │  │
                                │  └───────────────────────────┘  │
                                └─────────────────────────────────┘
```

## Project Structure

```
slidecue/
├── electron.vite.config.ts      # Vite bundler configuration
├── package.json
├── tsconfig.json
├── tailwind.config.js
├── icon/                        # Source icons
│   └── icon.png                 # Original app icon
├── resources/
│   ├── icon.png                 # Processed app icon
│   └── remote/
│       └── index.html           # Web remote UI (single file)
├── src/
│   ├── main/                    # Electron Main Process
│   │   ├── index.ts             # Entry point, IPC handlers
│   │   ├── pptx/
│   │   │   ├── slideParser.ts   # PPTX file parser (JSZip)
│   │   │   ├── macos.ts         # AppleScript automation
│   │   │   └── windows.ts       # COM automation (TODO)
│   │   └── server/
│   │       └── index.ts         # Express + Socket.IO server
│   ├── preload/
│   │   └── index.ts             # Context bridge (IPC)
│   └── renderer/                # Electron Renderer (Desktop UI)
│       ├── index.html
│       └── src/
│           ├── main.tsx
│           ├── App.tsx          # Main app component
│           ├── assets/
│           │   └── icon.png
│           ├── styles/
│           │   └── index.css
│           └── types/
│               └── electron.d.ts
└── out/                         # Build output
    ├── main/
    ├── preload/
    └── renderer/
```

## Component Details

### Main Process (`src/main/`)

#### `index.ts` — Application Entry Point
- Creates fullscreen `BrowserWindow`
- Registers IPC handlers for renderer communication
- Manages presentation lifecycle
- Cleans up temp files on quit

#### `pptx/slideParser.ts` — PPTX Parser
Extracts metadata from PowerPoint files using JSZip:
- Total slide count
- Hidden slide detection
- Speaker notes per slide (via relationship files)
- Slide dimensions

#### `pptx/macos.ts` — macOS Automation
Controls PowerPoint via AppleScript:
- `checkInstalled()` — Verify PowerPoint is installed
- `openPresentation()` — Open PPTX file
- `startSlideshow()` — Begin presentation mode
- `nextSlide()` / `prevSlide()` — Navigation
- `gotoSlide()` — Jump to specific slide (uses repeat loop workaround)
- `queryCurrentSlide()` — Poll current position
- Thumbnail generation via LibreOffice + pdftoppm

#### `server/index.ts` — Web Server
- Express serves remote UI and thumbnails
- Socket.IO with PIN authentication
- Auto-finds available port starting from 3000
- Broadcasts slide changes to all clients

### Preload Script (`src/preload/`)

Exposes safe APIs to renderer via `contextBridge`:
```typescript
electronAPI: {
  checkPowerPoint(): Promise<boolean>
  importPresentation(): Promise<PresentationInfo>
  startPresentation(): Promise<ConnectionInfo>
  stopPresentation(): Promise<void>
  onImportProgress(callback): () => void
}
```

### Renderer — Desktop UI (`src/renderer/`)

React application with states:
1. **checking** — Verifying PowerPoint installation
2. **no-powerpoint** — Error state if not installed
3. **idle** — Ready to import presentation
4. **importing** — Progress bar during thumbnail generation
5. **loaded** — Thumbnail grid, ready to present
6. **presenting** — QR code + PIN display

### Remote — Web UI (`resources/remote/`)

Single HTML file with embedded CSS/JS:
- PIN entry with auto-login via URL param
- Symmetric two-column layout (current + next slides)
- Speaker notes below each slide preview
- "Go to Slide" overview modal with thumbnails
- Timer/clock display in header
- Vibration feedback on mobile
- Keyboard shortcuts (arrows, space, G)
- Responsive design (stacks vertically on mobile)
- Clean dark theme without emojis

## Data Flow

### Import Flow
```
User clicks Import
        │
        ▼
File dialog ──► Select .pptx
        │
        ▼
Main Process:
  ├── Copy to temp directory
  ├── Parse PPTX (slides, notes, hidden)
  ├── Convert to PDF (LibreOffice)
  ├── Convert to PNGs (pdftoppm)
  └── Return metadata + thumbnail paths
        │
        ▼
Renderer shows thumbnail grid
```

### Present Flow
```
User clicks Present
        │
        ▼
Main Process:
  ├── Generate 4-digit PIN
  ├── Start Express server
  ├── Open presentation in PowerPoint
  └── Start slideshow mode
        │
        ▼
Return { url, pin }
        │
        ▼
Renderer shows QR code + PIN
```

### Remote Control Flow
```
Phone scans QR / enters PIN
        │
        ▼
Socket.IO connect with auth
        │
        ▼
Server validates PIN ──► Reject if invalid
        │
        ▼
Server sends current state
        │
        ▼
User taps Next
        │
        ▼
Socket emit('next')
        │
        ▼
Server calls macos.nextSlide()
        │
        ▼
Poll PowerPoint for new position
        │
        ▼
Broadcast('slide-changed', info)
        │
        ▼
All remotes update UI
```

## External Dependencies

### Runtime Requirements
- **Microsoft PowerPoint** — Required for presentation playback
- **LibreOffice** — Required for thumbnail generation (PPTX → PDF)
- **Poppler (pdftoppm)** — Required for thumbnail generation (PDF → PNG)

### macOS Installation
```bash
brew install --cask libreoffice
brew install poppler
```

## Security Model

### PIN Authentication
- Random 4-digit PIN per session
- Validated on Socket.IO handshake
- PIN only shown on presenter's screen
- QR code includes PIN for convenience

### Network Scope
- Server binds to `0.0.0.0` (local network only)
- No internet exposure by default
- Same-network devices can connect

## Tech Stack

| Layer | Technology |
|-------|------------|
| Framework | Electron 39 |
| Bundler | electron-vite + Vite 7 |
| UI | React 19 + TypeScript |
| Styling | Tailwind CSS 4 |
| Server | Express + Socket.IO |
| PPTX Parsing | JSZip + xml2js |
| QR Code | qrcode.react |
| Automation | AppleScript (macOS) |
