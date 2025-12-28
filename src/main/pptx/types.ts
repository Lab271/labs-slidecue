export interface SlideInfo {
  currentSlide: number;           // Actual PowerPoint slide number
  totalSlides: number;            // Total slides including hidden
  animationStep: number;          // Current animation step on this slide (0 = no animations played)
  animationsOnSlide: number;      // Total animations on current slide
  nextVisibleSlide: number | null; // Next slide that isn't hidden (for preview)
  isLastSlide: boolean;           // Are we on the last visible slide?
  currentNotes: string;           // Speaker notes for current slide
  nextNotes: string;              // Speaker notes for next visible slide
}

export interface SlideMetadata {
  thumbnails: string[];
  totalSlides: number;
  hiddenSlides: number[];
  visibleSlides: number[];
}

export type ProgressCallback = (current: number, total: number) => void;

export interface PowerPointAutomation {
  checkInstalled(): Promise<boolean>;
  openPresentation(filePath: string): Promise<void>;
  exportThumbnails(outputDir: string, onProgress?: ProgressCallback): Promise<SlideMetadata>;
  startSlideshow(): Promise<void>;
  nextSlide(): Promise<void>;
  prevSlide(): Promise<void>;
  gotoSlide(index: number): Promise<void>;
  getSlideInfo(): Promise<SlideInfo>;
  stopSlideshow(): Promise<void>;
  closePresentation(): Promise<void>;
}
