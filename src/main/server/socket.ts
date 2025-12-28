import { Server } from 'socket.io';
import { getAutomation } from '../pptx/automation';

export function setupSocketHandlers(io: Server) {
  const automation = getAutomation();

  io.on('connection', (socket) => {
    console.log('Remote connected:', socket.id);

    // Send current slide info on connect
    automation.getSlideInfo().then((info) => {
      socket.emit('slide-changed', info);
    });

    socket.on('next', async () => {
      try {
        await automation.nextSlide();
        const info = await automation.getSlideInfo();
        io.emit('slide-changed', info);
      } catch (error) {
        console.error('Error advancing slide:', error);
      }
    });

    socket.on('prev', async () => {
      try {
        await automation.prevSlide();
        const info = await automation.getSlideInfo();
        io.emit('slide-changed', info);
      } catch (error) {
        console.error('Error going to previous slide:', error);
      }
    });

    socket.on('goto', async (slideIndex: number) => {
      try {
        await automation.gotoSlide(slideIndex);
        const info = await automation.getSlideInfo();
        io.emit('slide-changed', info);
      } catch (error) {
        console.error('Error going to slide:', error);
      }
    });

    socket.on('disconnect', () => {
      console.log('Remote disconnected:', socket.id);
    });
  });
}
