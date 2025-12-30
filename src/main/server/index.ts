import express from 'express';
import { createServer } from 'http';
import { Server } from 'socket.io';
import { networkInterfaces } from 'os';
import { readdir } from 'fs/promises';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import log from 'electron-log';
import { app as electronApp } from 'electron';

let server: ReturnType<typeof createServer> | null = null;
let io: Server | null = null;
let currentPin: string = '';
let thumbnailsDir: string = '';
let currentPort: number = 3000;

function isPortAvailable(port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const testServer = createServer();
    testServer.once('error', (err) => {
      log.warn(`Port ${port} is not available:`, err.message);
      resolve(false);
    });
    testServer.once('listening', () => {
      testServer.close();
      log.info(`Port ${port} is available`);
      resolve(true);
    });
    testServer.listen(port, '0.0.0.0');
  });
}

async function findAvailablePort(startPort: number): Promise<number> {
  log.info(`Looking for available port starting from ${startPort}`);
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

export async function startServer(
  remoteUIPath: string,
  thumbsDir: string
): Promise<{
  url: string;
  pin: string;
  io: Server;
}> {
  log.info('Starting Express server...');
  log.info('Remote UI path:', remoteUIPath);
  log.info('Thumbnails directory:', thumbsDir);
  
  const app = express();
  server = createServer(app);
  io = new Server(server, {
    cors: { origin: '*' },
  });

  currentPin = generatePin();
  thumbnailsDir = thumbsDir;
  currentPort = await findAvailablePort(3000);
  
  log.info(`Found available port: ${currentPort}`);
  log.info(`Generated PIN: ${currentPin}`);

  // Serve remote UI at root
  app.use(express.static(remoteUIPath));

  // Serve the app icon
  app.get('/icon.png', (_req, res) => {
    // In packaged app, serve from resources; in dev, from project root
    const iconPath = electronApp.isPackaged 
      ? join(process.resourcesPath, 'icon.png')
      : join(process.cwd(), 'resources', 'icon.png');
    log.info('Serving icon from:', iconPath);
    res.sendFile(iconPath);
  });

  // Serve slide thumbnails
  app.use('/thumbnails', express.static(thumbnailsDir));
  
  // API to list available thumbnails
  app.get('/api/thumbnails', async (_req, res) => {
    try {
      const files = await readdir(thumbnailsDir);
      const images = files
        .filter(f => /\.(png|jpg|jpeg|gif)$/i.test(f))
        .sort()
        .map(f => `/thumbnails/${f}`);
      console.log('API thumbnails:', images);
      res.json({ thumbnails: images });
    } catch (e) {
      console.error('Error listing thumbnails:', e);
      res.json({ thumbnails: [] });
    }
  });

  // Socket.IO with PIN auth
  io.use((socket, next) => {
    const pin = socket.handshake.auth.pin;
    if (pin === currentPin) {
      next();
    } else {
      next(new Error('Invalid PIN'));
    }
  });

  server.listen(currentPort, '0.0.0.0', () => {
    log.info(`Server started successfully on port ${currentPort}`);
  });
  
  server.on('error', (err: any) => {
    log.error('Server error:', err);
    if (err.code === 'EADDRINUSE') {
      log.error(`Port ${currentPort} is already in use`);
    }
  });

  const localIP = getLocalIP();
  log.info(`Server URL: http://${localIP}:${currentPort}`);
  
  return {
    url: `http://${localIP}:${currentPort}`,
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
