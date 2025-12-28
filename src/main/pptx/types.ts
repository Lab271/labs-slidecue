export interface SlideInfo {
  current: number;
  total: number;
  hiddenSlides: number[];
}

export interface SlideMetadata {
  thumbnails: string[];
  totalSlides: number;
  hiddenSlides: number[];
  visibleSlideCount: number;
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
