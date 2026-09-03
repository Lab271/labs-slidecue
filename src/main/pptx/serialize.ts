// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { PowerPointAutomation } from './types';

/**
 * Wraps a PowerPoint backend so that at most one automation call is ever in
 * flight against it.
 *
 * Every backend keeps its slideshow position in module-level state
 * (`currentSlide`, `currentAnimationStep`, ...) and refreshes it from
 * PowerPoint across an `await` - an AppleScript round trip on macOS, a
 * PowerShell bridge command or a COM call on Windows. Nothing used to stop two
 * of those from overlapping: the web remote can fire `next` from several
 * devices at once, a fast double-tap queues two `nextSlide()` calls, and the
 * 500 ms poll in `server/socket.ts` calls `getSlideInfo()` on top of both. The
 * second caller would then read state that the first had not finished writing
 * and leave the cached position disagreeing with the real one.
 *
 * Chaining the calls onto a single promise fixes that at the source: the
 * commands reach PowerPoint in the order they were requested, and each one sees
 * the state the previous one committed. It also means a backend never has more
 * than one outstanding request to its transport, which is what the PowerShell
 * bridge's FIFO response matching already assumed.
 *
 * A rejected call does not poison the chain - the next queued call still runs,
 * and the rejection is delivered only to the caller that asked for it.
 */
export function serializeAutomation(backend: PowerPointAutomation): PowerPointAutomation {
  // Always settles fulfilled, so a failed call cannot stall everything behind it.
  let tail: Promise<void> = Promise.resolve();

  function run<T>(operation: () => Promise<T>): Promise<T> {
    const result = tail.then(operation);
    tail = result.then(
      () => undefined,
      () => undefined
    );
    return result;
  }

  // Listed one by one rather than proxied so that adding a method to
  // PowerPointAutomation without serializing it is a type error.
  return {
    checkInstalled: () => run(() => backend.checkInstalled()),
    openPresentation: (filePath) => run(() => backend.openPresentation(filePath)),
    exportThumbnails: (outputDir, onProgress) =>
      run(() => backend.exportThumbnails(outputDir, onProgress)),
    startSlideshow: () => run(() => backend.startSlideshow()),
    nextSlide: () => run(() => backend.nextSlide()),
    prevSlide: () => run(() => backend.prevSlide()),
    gotoSlide: (index) => run(() => backend.gotoSlide(index)),
    getSlideInfo: () => run(() => backend.getSlideInfo()),
    stopSlideshow: () => run(() => backend.stopSlideshow()),
    closePresentation: () => run(() => backend.closePresentation()),
  };
}
