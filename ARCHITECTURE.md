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
│   │   │   ├── serialize.ts     # One-command-at-a-time wrapper for a backend
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

## Packaging

SlideCue ships as an Electron installer (`.dmg`, `.zip`, `.exe`) built by
electron-builder. Two separate allowlists decide what ends up where, and they are
easy to confuse:

- **`build.files` in `package.json`** — read by *electron-builder*. Controls what goes
  into the app bundle (`out/**` plus `package.json`), with `build.extraResources`
  adding the web remote UI, the PowerShell bridge and the app icon alongside it.
- **`files` (top level) in `package.json`** — read by *npm*, for `npm pack` /
  `npm publish`. Controls the npm tarball, which the JFrog pilot publishes to
  Artifactory (see issue #56).

The npm allowlist is `out/**`, `resources/remote/**` and `resources/icon.png`. It is an
allowlist rather than a set of ignore rules because npm's default is "pack everything
not ignored", and for an application that default is wrong in both directions: it
shipped 3.8 MB of design-source icons while omitting things that matter. The three
entries are the paths the app resolves at runtime — `out/**` is the built main, preload
and renderer; `resources/remote/**` and `resources/icon.png` are what the Express
server serves when `app.isPackaged` is false (see `src/main/index.ts` and
`src/main/server/index.ts`).

Deliberately *not* in the npm tarball:

- `src/**` — the tarball ships built output; the source is on GitHub.
- `icon/**` — design source (five near-identical renders of the same icon, ~2.9 MB).
  It stays in the repository because `icon/icon_final.png` is the README banner, but
  nothing at runtime reads it.
- `resources/icon.icns`, `resources/icon.ico`, `resources/sbp.link.sp_space.png` —
  build-time inputs. The first two are consumed by electron-builder from a git checkout;
  the third is bundled into `out/renderer/` by Vite.

`prepack` runs `npm run build`, so `npm pack` cannot silently produce a tarball with no
application code in it — which is exactly what it did before, on any checkout where
`out/` had not been built.

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

#### `pptx/serialize.ts` — Command Serialization

Every backend caches the slideshow position in module-level state and refreshes
it from PowerPoint across an `await`. `serializeAutomation()` chains all calls
onto a single promise so only one is ever in flight, which is what keeps that
cache in step with PowerPoint when several remotes, a double-tap, and the 500 ms
`getSlideInfo()` poll all arrive at once. Each backend exports its automation
object already wrapped, so both `main/index.ts` and `server/socket.ts` share one
queue. Backend methods must therefore never call each other through the exported
object.

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
