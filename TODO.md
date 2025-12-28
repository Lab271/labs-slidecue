# TODO

## Phase 1: Project Setup
- [ ] Initialize Electron project with `electron-vite` (React + TypeScript)
- [ ] Install production dependencies (`express`, `socket.io`, `socket.io-client`, `bonjour-service`, `electron-updater`, `winax`)
- [ ] Install dev dependencies (`tailwindcss`, `postcss`, `autoprefixer`, types)
- [ ] Configure Tailwind CSS
- [ ] Configure Electron Forge for Windows builds (Squirrel maker)
- [ ] Create app icons (`icon.ico`, `icon.icns`, `icon.png`)

## Phase 2: PowerPoint Automation
- [ ] Create `src/main/pptx/types.ts` — TypeScript interfaces
- [ ] Create `src/main/pptx/macos.ts` — AppleScript implementation
- [ ] Create `src/main/pptx/windows.ts` — winax/COM implementation
- [ ] Create `src/main/pptx/automation.ts` — Platform abstraction
- [ ] Test `checkInstalled()` on macOS
- [ ] Test `openPresentation()` on macOS
- [ ] Test `exportThumbnails()` on macOS
- [ ] Test `startSlideshow()` on macOS
- [ ] Test `nextSlide()` / `prevSlide()` on macOS

## Phase 3: Web Server & Socket.IO
- [ ] Create `src/main/server/index.ts` — Express setup
- [ ] Implement `getLocalIP()` helper
- [ ] Implement `generatePin()` helper
- [ ] Implement `startServer()` with PIN auth middleware
- [ ] Create `src/main/server/socket.ts` — Socket.IO handlers
- [ ] Test server starts and serves static files
- [ ] Test PIN validation accepts correct PIN
- [ ] Test PIN validation rejects incorrect PIN

## Phase 4: Auto-Updater
- [ ] Create `src/main/updater.ts`
- [ ] Implement update check on startup
- [ ] Implement download prompt dialog
- [ ] Implement install on restart
- [ ] Test with a test release on GitHub

## Phase 5: Main Process
- [ ] Update `src/main/index.ts` entry point
- [ ] Implement `createWindow()` with preload
- [ ] Implement PowerPoint check on startup
- [ ] Implement IPC handler: `check-powerpoint`
- [ ] Implement IPC handler: `import-presentation`
- [ ] Implement IPC handler: `start-presentation`
- [ ] Implement IPC handler: `stop-presentation`
- [ ] Implement IPC handler: `get-slide-info`

## Phase 6: Preload Script
- [ ] Create/update `src/preload/index.ts`
- [ ] Expose `electronAPI` via contextBridge
- [ ] Add TypeScript global declarations

## Phase 7: Desktop UI (Renderer)
- [ ] Create `src/renderer/src/styles/index.css` with Tailwind
- [ ] Create `src/renderer/src/App.tsx` main component
- [ ] Implement "checking" state (loading spinner)
- [ ] Implement "no-powerpoint" state (error message)
- [ ] Implement "idle" state (import prompt)
- [ ] Implement "loaded" state (thumbnail grid + Present button)
- [ ] Implement "presenting" state (URL + PIN display)
- [ ] Style with Tailwind CSS

## Phase 8: Web Remote UI
- [ ] Create `src/remote/index.html`
- [ ] Create `src/remote/src/main.tsx`
- [ ] Create `src/remote/src/App.tsx`
- [ ] Create `src/remote/src/styles/index.css` with Tailwind
- [ ] Implement PIN entry form
- [ ] Implement Socket.IO connection with auth
- [ ] Implement slide preview (current/next toggle)
- [ ] Implement Previous/Next buttons
- [ ] Implement slide counter
- [ ] Make responsive for mobile screens
- [ ] Test on iPhone/Android browser

## Phase 9: CI/CD
- [ ] Create `.github/workflows/release.yml`
- [ ] Configure Windows build job
- [ ] Configure artifact upload
- [ ] Configure GitHub Release creation
- [ ] Test workflow with a `v0.1.0` tag
- [ ] Verify .exe downloads and installs

## Phase 10: Windows Testing
- [ ] Test on Windows with PowerPoint installed
- [ ] Verify `winax` COM automation works
- [ ] Test thumbnail export on Windows
- [ ] Test slideshow control on Windows
- [ ] Test web remote from phone on same network
- [ ] Test auto-updater from GitHub Release

## Polish & Extras
- [ ] Add loading states during thumbnail generation
- [ ] Add error handling for failed PowerPoint commands
- [ ] Add "Connecting..." state on remote
- [ ] Add connection lost / reconnecting UI
- [ ] Add presentation filename display
- [ ] Add keyboard shortcuts (arrow keys for slides)
- [ ] Add dark mode support
- [ ] Write user documentation
- [ ] Create demo video/GIF for README

## Future Ideas
- [ ] Speaker notes display on remote
- [ ] Timer/stopwatch on remote
- [ ] Laser pointer simulation
- [ ] Multiple presentation sessions
- [ ] Presenter view on secondary monitor
- [ ] Support for Keynote (macOS)
- [ ] Support for LibreOffice Impress (Linux)
