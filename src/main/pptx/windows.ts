import { PowerPointAutomation, SlideInfo } from './types';

// winax is Windows-only, conditionally import
let winax: any;
try {
  winax = require('winax');
} catch {
  // Not on Windows, will throw if methods are called
}

let pptApp: any = null;
let presentation: any = null;
let slideShow: any = null;

export const windowsAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      const testApp = new winax.Object('PowerPoint.Application');
      testApp.Quit();
      return true;
    } catch {
      return false;
    }
  },

  async openPresentation(filePath: string) {
    pptApp = new winax.Object('PowerPoint.Application');
    pptApp.Visible = true;
    presentation = pptApp.Presentations.Open(filePath);
  },

  async exportThumbnails(outputDir: string) {
    const paths: string[] = [];
    const count = presentation.Slides.Count;

    for (let i = 1; i <= count; i++) {
      const slide = presentation.Slides.Item(i);
      const filePath = `${outputDir}/slide_${i.toString().padStart(3, '0')}.png`;
      slide.Export(filePath, 'PNG', 1920, 1080);
      paths.push(filePath);
    }

    return paths;
  },

  async startSlideshow() {
    const settings = presentation.SlideShowSettings;
    settings.StartingSlide = 1;
    settings.EndingSlide = presentation.Slides.Count;
    slideShow = settings.Run();
  },

  async nextSlide() {
    if (slideShow?.View) {
      slideShow.View.Next();
    }
  },

  async prevSlide() {
    if (slideShow?.View) {
      slideShow.View.Previous();
    }
  },

  async gotoSlide(index: number) {
    if (slideShow?.View) {
      slideShow.View.GotoSlide(index);
    }
  },

  async getSlideInfo(): Promise<SlideInfo> {
    return {
      current: slideShow?.View?.CurrentShowPosition || 1,
      total: presentation?.Slides?.Count || 0,
    };
  },

  async stopSlideshow() {
    if (slideShow?.View) {
      slideShow.View.Exit();
      slideShow = null;
    }
  },

  async closePresentation() {
    if (presentation) {
      presentation.Close();
      presentation = null;
    }
    if (pptApp) {
      pptApp.Quit();
      pptApp = null;
    }
  },
};
