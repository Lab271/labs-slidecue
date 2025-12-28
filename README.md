# SlideCue

A PowerPoint presentation remote control app built with Electron.

**Control your PowerPoint presentations from any device on your network.**

![Status](https://img.shields.io/badge/status-in%20development-yellow)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- 📂 **Import PPTX** — Select any PowerPoint file and preview all slides as thumbnails
- 🎬 **Present with Full Animations** — Launches native PowerPoint in slideshow mode (all animations and transitions preserved)
- 📱 **Web Remote Control** — Any device on your local network can control the presentation
- 🔐 **PIN Security** — 4-digit PIN protects against unauthorized access
-  **Slide Preview** — See current or next slide on the remote device
- 🔄 **Auto-Updates** — App updates automatically from GitHub Releases

## Requirements

- **Microsoft PowerPoint** must be installed on the presentation machine
- **Windows 10/11** (deployment target)
- macOS supported for development

## Installation

### Download (Recommended)

Download the latest `.exe` installer from [GitHub Releases](https://github.com/yourusername/slidecue/releases).

### Build from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/slidecue.git
cd slidecue

# Install dependencies
npm install

# Run in development mode
npm run dev

# Build Windows installer
npm run make
```

## Usage

1. **Launch SlideCue** — The app checks if PowerPoint is installed
2. **Import a PPTX** — Click "Import" and select your PowerPoint file
3. **Preview Slides** — Thumbnails are generated for all slides
4. **Start Presenting** — Click "Present" to launch the slideshow
5. **Connect Remote** — Enter the URL shown on screen on your phone
6. **Enter PIN** — Type the 4-digit PIN shown on screen
7. **Control** — Use Previous/Next buttons to navigate slides

## Remote Interface

The web remote provides:

- **Previous / Next** buttons for slide navigation
- **Current / Next Slide** toggle to preview what's coming
- **Slide Counter** showing current position (e.g., "5 / 20")

## Technology Stack

- **Electron** — Cross-platform desktop app framework
- **React + TypeScript** — UI components
- **Tailwind CSS** — Styling
- **Express + Socket.IO** — Web server and real-time communication
- **winax** — Windows COM automation for PowerPoint control
- **AppleScript** — macOS PowerPoint control (development)
- **Electron Forge** — Packaging and distribution

## Development

```bash
# Start development server with hot reload
npm run dev

# Type check
npm run typecheck

# Lint
npm run lint

# Build for production
npm run build

# Create Windows installer
npm run make
```

## Architecture

See [ARCHITECTURE.md](./ARCHITECTURE.md) for technical details.

## Implementation Guide

See [IMPLEMENTATION.md](./IMPLEMENTATION.md) for step-by-step build instructions.

## Roadmap

See [TODO.md](./TODO.md) for the current task list.

## Contributing

Contributions are welcome! Please read the implementation guide first.

## License

MIT License — see [LICENSE](./LICENSE) for details.
