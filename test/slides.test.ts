import { describe, it, expect } from 'vitest';
import {
  visibleToActualIndex,
  actualToVisibleIndex,
} from '../src/main/pptx/parseHiddenSlides';
import {
  getNextVisibleSlide,
  getPrevVisibleSlide,
  getSlideData,
  type PresentationData,
} from '../src/main/pptx/slideParser';

describe('hidden-slide index mapping', () => {
  it('is an identity when no slides are hidden', () => {
    for (let i = 0; i <= 5; i++) {
      expect(visibleToActualIndex(i, [])).toBe(i);
      expect(actualToVisibleIndex(i, [])).toBe(i);
    }
  });

  it('skips hidden slides when mapping visible → actual', () => {
    // slide 2 hidden: 1st visible = 1, 2nd visible = 3, 3rd visible = 4
    const hidden = [2];
    expect(visibleToActualIndex(1, hidden)).toBe(1);
    expect(visibleToActualIndex(2, hidden)).toBe(3);
    expect(visibleToActualIndex(3, hidden)).toBe(4);
  });

  it('counts only visible slides when mapping actual → visible', () => {
    const hidden = [2];
    expect(actualToVisibleIndex(1, hidden)).toBe(1);
    expect(actualToVisibleIndex(2, hidden)).toBe(1); // hidden slide doesn't add
    expect(actualToVisibleIndex(3, hidden)).toBe(2);
  });

  it('round-trips visible → actual → visible for several hidden layouts', () => {
    for (const hidden of [[], [2], [1, 4], [3, 5, 6]]) {
      for (let visible = 1; visible <= 4; visible++) {
        const actual = visibleToActualIndex(visible, hidden);
        expect(actualToVisibleIndex(actual, hidden)).toBe(visible);
      }
    }
  });
});

describe('slide navigation', () => {
  const data: PresentationData = {
    slides: [
      { slideNumber: 1, name: 'a', hidden: false, animationClicks: 0, notes: '' },
      { slideNumber: 2, name: 'b', hidden: true, animationClicks: 0, notes: 'secret' },
      { slideNumber: 3, name: 'c', hidden: false, animationClicks: 2, notes: '' },
    ],
    totalSlides: 3,
    visibleSlides: [1, 3],
    hiddenSlides: [2],
  };

  it('advances to the next visible slide, skipping hidden ones', () => {
    expect(getNextVisibleSlide(1, data)).toBe(3);
    expect(getNextVisibleSlide(3, data)).toBeNull(); // nothing after the last
  });

  it('goes back to the previous visible slide', () => {
    expect(getPrevVisibleSlide(3, data)).toBe(1);
    expect(getPrevVisibleSlide(1, data)).toBeNull(); // nothing before the first
  });

  it('looks up slide data by number and returns null when absent', () => {
    expect(getSlideData(2, data)?.notes).toBe('secret');
    expect(getSlideData(99, data)).toBeNull();
  });
});
