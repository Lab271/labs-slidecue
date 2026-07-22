// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
// Simple test server for remote UI development
// Run with: node test-remote.js
// Then open: http://localhost:3333?demo=true

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3333;

// Mock slide data
const mockSlideInfo = {
  currentSlide: 3,
  totalSlides: 12,
  nextVisibleSlide: 4,
  animationStep: 1,
  animationsOnSlide: 3,
  isLastSlide: false,
  currentNotes: 'These are the speaker notes for the current slide. You can add multiple lines of notes here to test how they display and scroll.\n\nThis is a second paragraph to test spacing.',
  nextNotes: 'Notes for the upcoming slide go here.'
};

const server = http.createServer((req, res) => {
  let filePath = req.url.split('?')[0];
  
  if (filePath === '/' || filePath === '/index.html') {
    filePath = '/index.html';
  }
  
  // Serve mock thumbnails (gray placeholder images)
  if (filePath.startsWith('/thumbnails/')) {
    res.writeHead(200, { 'Content-Type': 'image/svg+xml' });
    const slideNum = filePath.match(/slide_(\d+)/)?.[1] || '1';
    res.end(`<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080">
      <rect fill="#1e293b" width="1920" height="1080"/>
      <text x="960" y="540" text-anchor="middle" fill="#64748b" font-family="system-ui" font-size="120" font-weight="bold">Slide ${parseInt(slideNum)}</text>
    </svg>`);
    return;
  }
  
  // Serve icon
  if (filePath === '/icon.png') {
    const iconPath = path.join(__dirname, 'resources', 'icon.png');
    if (fs.existsSync(iconPath)) {
      res.writeHead(200, { 'Content-Type': 'image/png' });
      fs.createReadStream(iconPath).pipe(res);
      return;
    }
    res.writeHead(404);
    res.end();
    return;
  }
  
  // Serve socket.io mock
  if (filePath === '/socket.io/socket.io.js') {
    res.writeHead(200, { 'Content-Type': 'application/javascript' });
    res.end(`
      // Mock socket.io for testing
      window.io = function(url, opts) {
        const handlers = {};
        const socket = {
          on: (event, fn) => { handlers[event] = fn; },
          emit: (event, data) => {
            console.log('Socket emit:', event, data);
            if (event === 'next') {
              window.mockSlideInfo.currentSlide = Math.min(window.mockSlideInfo.currentSlide + 1, window.mockSlideInfo.totalSlides);
              window.mockSlideInfo.nextVisibleSlide = window.mockSlideInfo.currentSlide + 1;
              window.mockSlideInfo.isLastSlide = window.mockSlideInfo.currentSlide >= window.mockSlideInfo.totalSlides;
              handlers['slide-changed']?.(window.mockSlideInfo);
            } else if (event === 'prev') {
              window.mockSlideInfo.currentSlide = Math.max(window.mockSlideInfo.currentSlide - 1, 1);
              window.mockSlideInfo.nextVisibleSlide = window.mockSlideInfo.currentSlide + 1;
              window.mockSlideInfo.isLastSlide = false;
              handlers['slide-changed']?.(window.mockSlideInfo);
            } else if (event === 'goto') {
              window.mockSlideInfo.currentSlide = data;
              window.mockSlideInfo.nextVisibleSlide = data + 1;
              window.mockSlideInfo.isLastSlide = data >= window.mockSlideInfo.totalSlides;
              handlers['slide-changed']?.(window.mockSlideInfo);
            }
          },
          close: () => {}
        };
        
        // Auto-connect after short delay
        setTimeout(() => {
          handlers['connect']?.();
          handlers['slide-changed']?.(window.mockSlideInfo);
        }, 100);
        
        return socket;
      };
      
      window.mockSlideInfo = ${JSON.stringify(mockSlideInfo)};
    `);
    return;
  }
  
  // Serve the remote HTML
  const fullPath = path.join(__dirname, 'resources', 'remote', filePath);
  
  if (fs.existsSync(fullPath)) {
    const ext = path.extname(fullPath);
    const contentTypes = {
      '.html': 'text/html',
      '.css': 'text/css',
      '.js': 'application/javascript',
      '.png': 'image/png',
      '.svg': 'image/svg+xml'
    };
    res.writeHead(200, { 'Content-Type': contentTypes[ext] || 'text/plain' });
    fs.createReadStream(fullPath).pipe(res);
  } else {
    res.writeHead(404);
    res.end('Not found');
  }
});

server.listen(PORT, () => {
  console.log(`
┌─────────────────────────────────────────────────┐
│  Remote UI Test Server                          │
├─────────────────────────────────────────────────┤
│  URL: http://localhost:${PORT}?pin=1234            │
│                                                 │
│  • Enter any 4-digit PIN to connect             │
│  • Mock slides will be shown                    │
│  • Prev/Next buttons work                       │
│  • Test responsive design in DevTools           │
│                                                 │
│  Press Ctrl+C to stop                           │
└─────────────────────────────────────────────────┘
`);
});
