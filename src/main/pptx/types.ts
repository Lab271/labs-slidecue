export interface SlideInfo {
  current: number;
  total: number;
}

export interface PowerPointAutomation {
  checkInstalled(): Promise<boolean>;
  openPresentation(filePath: string): Promise<void>;
  exportThumbnails(outputDir: string): Promise<string[]>;
  startSlideshow(): Promise<void>;
  nextSlide(): Promise<void>;
  prevSlide(): Promise<void>;
  gotoSlide(index: number): Promise<void>;
  getSlideInfo(): Promise<SlideInfo>;
  stopSlideshow(): Promise<void>;
  closePresentation(): Promise<void>;
}
