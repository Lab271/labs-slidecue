# SlideCue

Control PowerPoint presentations from any device on your network.

<p align="center">
  <img src="icon/icon_final.png" alt="SlideCue" width="200" />
</p>

## Features

- **Remote Control** — Control your presentation from your phone or tablet
- **QR Code Login** — Scan to connect instantly, no typing required
- **Speaker Notes** — View notes for current and upcoming slides
- **Timer** — Track elapsed time during your presentation
- **PIN Protected** — Secure 4-digit PIN for each session
- **Dark Theme** — Clean, professional dark UI throughout
- **Responsive Design** — Works on phones, tablets, and laptops

## Requirements

### macOS
- Microsoft PowerPoint (tested with Microsoft 365)
- LibreOffice (for thumbnail generation)
- Poppler (for PDF to image conversion)

```bash
# Install dependencies
brew install --cask libreoffice
brew install poppler
```

### Windows
- Microsoft PowerPoint (Microsoft 365 or standalone)
- **PowerShell Bridge**: Uses PowerShell COM automation for PowerPoint control (no additional dependencies needed)

## Troubleshooting

### Logs Location
If you encounter issues, check the application logs:
- **macOS**: `~/Library/Logs/SlideCue/main.log`
- **Windows**: `%USERPROFILE%\AppData\Roaming\SlideCue\logs\main.log`

### Common Issues
- **PNGs not generating**: Ensure LibreOffice and Poppler are installed on macOS
- **Server won't start**: Check if port 3000 is available or being blocked by firewall
- **Permission errors**: Grant Accessibility and Automation permissions to SlideCue in System Settings

## Getting Started

### Development

```bash
# Install dependencies
npm install

# Run in development mode
npm run dev

# Build for production
npm run build

# Start production build
npm start
```

### Usage

1. **Import** — Click "Import PPTX" and select your presentation
2. **Wait** — Thumbnails are generated (may take a moment for large files)
3. **Present** — Click "Present" to start the slideshow
4. **Connect** — Scan the QR code or enter the URL and PIN on your phone
5. **Control** — Use the remote to navigate slides

## Remote Controls

| Action | Button | Keyboard |
|--------|--------|----------|
| Next slide/animation | Next → | Arrow Right, Space |
| Previous slide | ← Previous | Arrow Left |
| Go to specific slide | Go to Slide | G |
| Close overview | Close | Escape |

## How It Works

1. **PPTX Parsing** — Extracts slide count, notes, and hidden slide info
2. **Thumbnail Generation** — LibreOffice converts to PDF, pdftoppm creates PNGs
3. **PowerPoint Control** — AppleScript commands control the native app
4. **Web Server** — Express serves the remote UI on your local network
5. **Real-time Sync** — Socket.IO keeps all connected devices in sync

## Project Structure

```
src/
├── main/           # Electron main process
│   ├── pptx/       # PowerPoint automation
│   └── server/     # Express + Socket.IO
├── preload/        # IPC bridge
└── renderer/       # Desktop UI (React)

resources/
└── remote/         # Web remote UI
```

## Tech Stack

- **Electron** — Desktop application framework
- **React** — UI components
- **Tailwind CSS** — Styling
- **Socket.IO** — Real-time communication
- **JSZip** — PPTX file parsing
- **AppleScript** — PowerPoint automation (macOS)

## Roadmap

- [x] Windows support (PowerShell COM automation)
- [x] Packaged releases (.app, .exe)
- [ ] Auto-updater
- [ ] Laser pointer simulation

## License

Licensed under the [Apache License 2.0](LICENSE). Copyright © 2025-2026 Schuberg Philis / Lab271. See [NOTICE](NOTICE).
