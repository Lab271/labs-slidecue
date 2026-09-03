// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { Server } from 'socket.io';
import { getAutomation } from '../pptx/automation';
import { PowerPointAutomation } from '../pptx/types';
import log from 'electron-log';

let pollInterval: NodeJS.Timeout | null = null;
let pollInFlight: Promise<void> | null = null;
let lastSlideNumber = 1;

export function setupSocketHandlers(io: Server) {
  const automation = getAutomation();

  // Start polling for slide changes
  startSlidePolling(io, automation);

  io.on('connection', (socket) => {
    log.info('Remote connected:', socket.id);

    // Send current slide info on connect
    automation.getSlideInfo().then((info) => {
      lastSlideNumber = info.currentSlide;
      socket.emit('slide-changed', info);
    });

    socket.on('next', async () => {
      try {
        await automation.nextSlide();
        const info = await automation.getSlideInfo();
        lastSlideNumber = info.currentSlide;
        io.emit('slide-changed', info);
      } catch (error) {
        log.error('Error advancing slide:', error);
      }
    });

    socket.on('prev', async () => {
      try {
        await automation.prevSlide();
        const info = await automation.getSlideInfo();
        lastSlideNumber = info.currentSlide;
        io.emit('slide-changed', info);
      } catch (error) {
        log.error('Error going to previous slide:', error);
      }
    });

    socket.on('goto', async (slideIndex: number) => {
      try {
        await automation.gotoSlide(slideIndex);
        const info = await automation.getSlideInfo();
        lastSlideNumber = info.currentSlide;
        io.emit('slide-changed', info);
      } catch (error) {
        log.error('Error going to slide:', error);
      }
    });

    socket.on('disconnect', () => {
      log.info('Remote disconnected:', socket.id);
    });
  });
}
function startSlidePolling(io: Server, automation: PowerPointAutomation) {
  // Stop existing poll if any
  if (pollInterval) {
    clearInterval(pollInterval);
  }

  // Poll every 500ms for slide changes. Automation calls are serialized per
  // backend, so a poll that outlives its tick would queue behind the next one
  // and behind every remote command instead of overlapping with them - skip the
  // tick rather than let the backlog grow.
  pollInterval = setInterval(() => {
    if (pollInFlight) {
      return;
    }
    pollInFlight = pollOnce(io, automation).finally(() => {
      pollInFlight = null;
    });
  }, 500);
}

async function pollOnce(io: Server, automation: PowerPointAutomation): Promise<void> {
  try {
    const info = await automation.getSlideInfo();

    // Only emit if slide actually changed
    if (info.currentSlide !== lastSlideNumber) {
      log.info(`Slide changed from ${lastSlideNumber} to ${info.currentSlide}`);
      lastSlideNumber = info.currentSlide;
      io.emit('slide-changed', info);
    }
  } catch {
    // Ignore errors during polling
  }
}

export function stopSlidePolling() {
  if (pollInterval) {
    clearInterval(pollInterval);
    pollInterval = null;
  }
}
