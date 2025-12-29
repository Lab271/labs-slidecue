import { Server } from 'socket.io';
import { getAutomation } from '../pptx/automation';
import log from 'electron-log';

let pollInterval: NodeJS.Timeout | null = null;
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
function startSlidePolling(io: Server, automation: any) {
  // Stop existing poll if any
  if (pollInterval) {
    clearInterval(pollInterval);
  }

  // Poll every 500ms for slide changes
  pollInterval = setInterval(async () => {
    try {
      const info = await automation.getSlideInfo();
      
      // Only emit if slide actually changed
      if (info.currentSlide !== lastSlideNumber) {
        log.info(`Slide changed from ${lastSlideNumber} to ${info.currentSlide}`);
        lastSlideNumber = info.currentSlide;
        io.emit('slide-changed', info);
      }
    } catch (error) {
      // Ignore errors during polling
    }
  }, 500);
}

export function stopSlidePolling() {
  if (pollInterval) {
    clearInterval(pollInterval);
    pollInterval = null;
  }
}
