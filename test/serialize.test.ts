// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { describe, it, expect } from 'vitest';
import { serializeAutomation } from '../src/main/pptx/serialize';
import type { PowerPointAutomation, SlideInfo, SlideMetadata } from '../src/main/pptx/types';

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));

const emptyInfo: SlideInfo = {
  currentSlide: 1,
  totalSlides: 1,
  animationStep: 0,
  animationsOnSlide: 0,
  nextVisibleSlide: null,
  isLastSlide: true,
  currentNotes: '',
  nextNotes: '',
};

const emptyMetadata: SlideMetadata = {
  thumbnails: [],
  totalSlides: 1,
  hiddenSlides: [],
  visibleSlides: [],
};

/** A backend whose methods all resolve, for tests that only care about one of them. */
function stubBackend(overrides: Partial<PowerPointAutomation> = {}): PowerPointAutomation {
  return {
    checkInstalled: async () => true,
    openPresentation: async () => {},
    exportThumbnails: async () => emptyMetadata,
    startSlideshow: async () => {},
    nextSlide: async () => {},
    prevSlide: async () => {},
    gotoSlide: async () => {},
    getSlideInfo: async () => emptyInfo,
    stopSlideshow: async () => {},
    closePresentation: async () => {},
    ...overrides,
  };
}

/**
 * Stands in for a PowerPoint backend: `devicePosition` is what PowerPoint
 * itself is showing, `currentSlide` is the cached copy the backend keeps. The
 * cached value is read before the first await and written after the second,
 * which is the shape every real backend has.
 */
function fakePowerPoint() {
  const seen = { maxConcurrent: 0, inFlight: 0, order: [] as string[] };
  let devicePosition = 1;
  let currentSlide = 1;

  async function command(label: string, move: (from: number) => number): Promise<void> {
    seen.inFlight += 1;
    seen.maxConcurrent = Math.max(seen.maxConcurrent, seen.inFlight);
    seen.order.push(label);
    try {
      const before = currentSlide;
      await tick(); // round trip out to PowerPoint
      devicePosition = move(before);
      await tick(); // round trip back
      // Reproducing the hazard on purpose - this is the shape serializeAutomation
      // exists to make safe, so the rule that bans it elsewhere is off here.
      // eslint-disable-next-line require-atomic-updates
      currentSlide = devicePosition;
    } finally {
      seen.inFlight -= 1;
    }
  }

  const backend = stubBackend({
    nextSlide: () => command('next', (from) => from + 1),
    prevSlide: () => command('prev', (from) => Math.max(1, from - 1)),
    getSlideInfo: async () => {
      await tick();
      return { ...emptyInfo, currentSlide };
    },
  });

  return { backend, seen, position: () => currentSlide };
}

describe('serializeAutomation', () => {
  it('lets an unserialized backend lose an update (the bug being fixed)', async () => {
    const { backend } = fakePowerPoint();

    await Promise.all([backend.nextSlide(), backend.nextSlide()]);

    // Both calls read slide 1 before their await, so both advanced to 2.
    expect((await backend.getSlideInfo()).currentSlide).toBe(2);
  });

  it('applies concurrent commands one at a time', async () => {
    const { backend, seen } = fakePowerPoint();
    const automation = serializeAutomation(backend);

    await Promise.all([automation.nextSlide(), automation.nextSlide()]);

    expect(seen.maxConcurrent).toBe(1);
    expect((await automation.getSlideInfo()).currentSlide).toBe(3);
  });

  it('runs commands in the order they were requested', async () => {
    const { backend, seen } = fakePowerPoint();
    const automation = serializeAutomation(backend);

    await Promise.all([
      automation.nextSlide(),
      automation.nextSlide(),
      automation.prevSlide(),
      automation.nextSlide(),
    ]);

    expect(seen.order).toEqual(['next', 'next', 'prev', 'next']);
    expect((await automation.getSlideInfo()).currentSlide).toBe(3);
  });

  it('keeps a polling read from interleaving with a command', async () => {
    const { backend, seen } = fakePowerPoint();
    const automation = serializeAutomation(backend);

    const [, info] = await Promise.all([automation.nextSlide(), automation.getSlideInfo()]);

    expect(seen.maxConcurrent).toBe(1);
    // The poll was queued behind the command, so it reports the settled position.
    expect(info.currentSlide).toBe(2);
  });

  it('delivers a rejection only to its own caller and keeps the queue moving', async () => {
    const calls: string[] = [];
    const automation = serializeAutomation(
      stubBackend({
        startSlideshow: async () => {
          calls.push('start');
          throw new Error('PowerPoint said no');
        },
        nextSlide: async () => {
          calls.push('next');
        },
      })
    );

    const failing = automation.startSlideshow();
    const following = automation.nextSlide();

    await expect(failing).rejects.toThrow('PowerPoint said no');
    await expect(following).resolves.toBeUndefined();
    expect(calls).toEqual(['start', 'next']);
  });

  it('forwards arguments and return values', async () => {
    const opened: string[] = [];
    const progress: Array<[number, number]> = [];
    const metadata: SlideMetadata = { ...emptyMetadata, totalSlides: 7 };

    const automation = serializeAutomation(
      stubBackend({
        openPresentation: async (filePath) => {
          opened.push(filePath);
        },
        exportThumbnails: async (outputDir, onProgress) => {
          onProgress?.(1, 2);
          return { ...metadata, thumbnails: [`${outputDir}/slide_001.png`] };
        },
        gotoSlide: async (index) => {
          opened.push(`goto:${index}`);
        },
      })
    );

    await automation.openPresentation('/tmp/deck.pptx');
    await automation.gotoSlide(4);
    const result = await automation.exportThumbnails('/tmp/thumbs', (current, total) =>
      progress.push([current, total])
    );

    expect(opened).toEqual(['/tmp/deck.pptx', 'goto:4']);
    expect(progress).toEqual([[1, 2]]);
    expect(result.totalSlides).toBe(7);
    expect(result.thumbnails).toEqual(['/tmp/thumbs/slide_001.png']);
  });
});
