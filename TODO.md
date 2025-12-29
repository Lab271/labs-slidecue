# TODO

## ✅ Phase 1: Project Setup — COMPLETE
- [x] Initialize Electron project with `electron-vite` (React + TypeScript)
- [x] Install production dependencies (`express`, `socket.io`, `qrcode.react`)
- [x] Install dev dependencies (`tailwindcss`, `postcss`, `autoprefixer`, types)
- [x] Configure Tailwind CSS v4
- [x] Create app icon (`icon.png`)

## ✅ Phase 2: PowerPoint Automation (macOS) — COMPLETE
- [x] Create `src/main/pptx/slideParser.ts` — PPTX parsing with JSZip
- [x] Create `src/main/pptx/macos.ts` — AppleScript implementation
- [x] Implement `checkInstalled()` — checks for PowerPoint.app
- [x] Implement `openPresentation()` — opens PPTX in PowerPoint
- [x] Implement thumbnail generation via LibreOffice + pdftoppm
- [x] Implement `startSlideshow()` — starts presentation mode
- [x] Implement `nextSlide()` / `prevSlide()` — navigation
- [x] Implement `gotoSlide()` — jump to specific slide
- [x] Implement speaker notes extraction from PPTX
- [x] Handle hidden slides detection

## ✅ Phase 3: Web Server & Socket.IO — COMPLETE
- [x] Create `src/main/server/index.ts` — Express setup
- [x] Implement `getLocalIP()` helper
- [x] Implement `generatePin()` helper
- [x] Implement `startServer()` with PIN auth middleware
- [x] Socket.IO handlers for `next`, `prev`, `goto`
- [x] Broadcast slide changes to all connected clients
- [x] Serve thumbnails and app icon

## ✅ Phase 5: Main Process — COMPLETE
- [x] `src/main/index.ts` entry point
- [x] `createWindow()` with fullscreen support
- [x] IPC handler: `check-powerpoint`
- [x] IPC handler: `import-presentation`
- [x] IPC handler: `start-presentation`
- [x] IPC handler: `stop-presentation`
- [x] Import progress reporting

## ✅ Phase 6: Preload Script — COMPLETE
- [x] `src/preload/index.ts`
- [x] Expose `electronAPI` via contextBridge
- [x] TypeScript global declarations

## ✅ Phase 7: Desktop UI (Renderer) — COMPLETE
- [x] Dark theme design throughout
- [x] "checking" state with animated logo
- [x] "no-powerpoint" state with error message
- [x] "idle" state with import dropzone
- [x] "importing" state with progress bar and steps
- [x] "loaded" state with thumbnail grid
- [x] "presenting" state with QR code and PIN display
- [x] Responsive grid layout for thumbnails

## ✅ Phase 8: Web Remote UI — COMPLETE
- [x] `resources/remote/index.html` — Single-file remote UI
- [x] PIN entry form with auto-focus
- [x] Auto-login via URL parameter (`?pin=1234`)
- [x] Two-column layout (current + next slide)
- [x] Speaker notes display for current and next slides
- [x] Previous/Next navigation buttons
- [x] "Go to Slide" overview modal
- [x] Slide counter with animation tracking
- [x] Elapsed timer + current time display
- [x] Vibration feedback on button press
- [x] Keyboard shortcuts (arrow keys, G for overview)
- [x] Dark theme, mobile-optimized design

## ⏳ Phase 4: Auto-Updater — NOT STARTED
- [ ] Create `src/main/updater.ts`
- [ ] Implement update check on startup
- [ ] Implement download prompt dialog
- [ ] Implement install on restart
- [ ] Test with a test release on GitHub

## ⏳ Phase 9: CI/CD — NOT STARTED
- [ ] Create `.github/workflows/release.yml`
- [ ] Configure macOS build job
- [ ] Configure Windows build job
- [ ] Configure artifact upload
- [ ] Configure GitHub Release creation
- [ ] Test workflow with a `v0.1.0` tag

## ✅ Phase 10: Windows Support — COMPLETE
- [x] Create `src/main/pptx/windows.ts` — PowerShell COM implementation
- [x] Test on Windows with PowerPoint installed
- [x] Verify COM automation works
- [x] Test thumbnail export on Windows
- [x] Test slideshow control on Windows

## ✅ Phase 11: Packaging — COMPLETE
- [x] Configure electron-forge for macOS (.app, .dmg)
- [x] Configure electron-forge for Windows (.exe)
- [x] Create `icon.icns` for macOS
- [x] Create `icon.ico` for Windows
- [x] Test packaged app on macOS
- [x] Test packaged app on Windows

## Polish & Extras
- [x] Loading states during thumbnail generation
- [x] Error handling for failed PowerPoint commands
- [x] Connection info display with QR code
- [x] Presentation filename display
- [x] Keyboard shortcuts (arrow keys for slides)
- [x] Dark mode throughout
- [x] Speaker notes on remote
- [x] Timer/clock on remote
- [x] Haptic feedback on remote
- [x] Symmetric two-column remote layout
- [x] Responsive mobile design for remote
- [x] Removed emojis for cleaner look
- [x] Consistent design between main app and remote
- [ ] Connection lost / reconnecting UI on remote
- [ ] Write user documentation
- [ ] Create demo video/GIF for README

## Future Ideas
- [ ] Laser pointer simulation
- [ ] Multiple simultaneous presenters
- [ ] Presenter view on secondary monitor
- [ ] Support for Keynote (macOS)
- [ ] Support for Google Slides (web)
- [ ] Cloud sync for presentations
