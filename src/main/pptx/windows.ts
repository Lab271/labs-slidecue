import { exec } from 'child_process';
import { promisify } from 'util';
import { readdir, mkdir, unlink, copyFile } from 'fs/promises';
import { join, basename } from 'path';
import { tmpdir } from 'os';
import { PowerPointAutomation, SlideInfo, SlideMetadata, ProgressCallback } from './types';
import { parsePresentationData, PresentationData, getNextVisibleSlide, getSlideData } from './slideParser';
import log from 'electron-log';

const execAsync = promisify(exec);

// PowerShell-based COM automation (no winax needed)
async function runPowerShell(script: string): Promise<string> {
  const { stdout } = await execAsync(`powershell -NoProfile -Command "${script.replace(/"/g, '\\"')}"`);
  return stdout.trim();
}

// Presentation state
let pptApp: any = null;
let presentation: any = null;
let slideShow: any = null;
let presentationData: PresentationData | null = null;
let currentSlide = 1;
let currentAnimationStep = 0;
let totalSlides = 1;
let currentPresentationPath = '';
let localPresentationCopy = '';

/**
 * Query PowerPoint for the actual current slide number
 */
function queryCurrentSlide(): number {
  try {
    if (slideShow?.View) {
      return slideShow.View.CurrentShowPosition || 1;
    }
  } catch {
    // Ignore
  }
  return currentSlide;
}

export const windowsAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      await runPowerShell('$ppt = New-Object -ComObject PowerPoint.Application; $ppt.Version; $ppt.Quit()');
      return true;
    } catch {
      return false;
    }
  },
      return false;
    }
  },

  async openPresentation(filePath: string) {
    currentPresentationPath = filePath;
    
    // Copy file to local temp directory to avoid permission issues with cloud storage
    const tempDir = join(tmpdir(), 'slidecue-presentations');
    await mkdir(tempDir, { recursive: true });
    localPresentationCopy = join(tempDir, basename(filePath));
    
    console.log('Copying presentation to temp location...');
    await copyFile(filePath, localPresentationCopy);
    
    // Parse presentation data (slides, hidden, animations, notes)
    console.log('Parsing presentation data...');
    presentationData = await parsePresentationData(localPresentationCopy);
    
    console.log('Presentation data:', JSON.stringify(presentationData, null, 2));
    
    // Open in PowerPoint via COM
    console.log('Opening presentation...');
    pptApp = new winax.Object('PowerPoint.Application');
    pptApp.Visible = true;
    presentation = pptApp.Presentations.Open(localPresentationCopy);
    
    // Get total slides
    totalSlides = presentation.Slides.Count || presentationData.totalSlides;
    currentSlide = 1;
    currentAnimationStep = 0;
    
    console.log(`Opened presentation with ${totalSlides} slides`);
    console.log(`Visible slides: ${presentationData.visibleSlides.join(', ')}`);
    console.log(`Hidden slides: ${presentationData.hiddenSlides.join(', ')}`);
  },

  async exportThumbnails(outputDir: string, onProgress?: ProgressCallback): Promise<SlideMetadata> {
    await mkdir(outputDir, { recursive: true });
    
    const pptxPath = localPresentationCopy || currentPresentationPath;
    const paths: string[] = [];
    
    console.log('Exporting thumbnails...');
    onProgress?.(1, totalSlides);
    
    // Use visible slides from parsed data
    const visibleSlides = presentationData?.visibleSlides || [];
    
    // Export only visible slides
    for (let i = 0; i < visibleSlides.length; i++) {
      const slideNum = visibleSlides[i];
      const slide = presentation.Slides.Item(slideNum);
      const filePath = join(outputDir, `slide_${String(slideNum).padStart(3, '0')}.png`);
      
      try {
        slide.Export(filePath, 'PNG', 1920, 1080);
        paths.push(filePath);
        console.log(`Exported slide ${slideNum}`);
      } catch (e) {
        console.error(`Failed to export slide ${slideNum}:`, e);
      }
      
      onProgress?.(i + 1, visibleSlides.length);
    }
    
    console.log('Generated thumbnails:', paths);
    
    return {
      thumbnails: paths,
      totalSlides,
      hiddenSlides: presentationData?.hiddenSlides || [],
      visibleSlides: presentationData?.visibleSlides || [],
    };
  },

  async startSlideshow() {
    currentSlide = 1;
    currentAnimationStep = 0;
    
    const settings = presentation.SlideShowSettings;
    settings.StartingSlide = 1;
    settings.EndingSlide = totalSlides;
    slideShow = settings.Run();
    
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // PowerPoint may skip hidden slide 1, query actual position
    currentSlide = queryCurrentSlide();
  },

  async nextSlide() {
    const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    
    if (slideShow?.View) {
      slideShow.View.Next();
    }
    
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Check if we advanced an animation or moved to next slide
    if (currentAnimationStep < animationsOnSlide) {
      currentAnimationStep++;
    }
    
    const actualSlide = queryCurrentSlide();
    
    if (actualSlide !== currentSlide) {
      currentSlide = actualSlide;
      currentAnimationStep = 0;
      console.log(`Moved to slide ${currentSlide}`);
    } else {
      console.log(`Animation ${currentAnimationStep}/${animationsOnSlide} on slide ${currentSlide}`);
    }
  },

  async prevSlide() {
    if (slideShow?.View) {
      slideShow.View.Previous();
    }
    
    await new Promise(resolve => setTimeout(resolve, 100));
    
    currentSlide = queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Moved to slide ${currentSlide}`);
  },

  async gotoSlide(slideNumber: number) {
    console.log(`Going to slide ${slideNumber}...`);
    
    if (slideShow?.View) {
      try {
        // Windows COM should support GotoSlide directly
        slideShow.View.GotoSlide(slideNumber);
      } catch (e) {
        console.error('GotoSlide failed, trying navigation approach:', e);
        // Fallback: navigate via First + Next
        slideShow.View.First();
        await new Promise(resolve => setTimeout(resolve, 50));
        
        let pos = queryCurrentSlide();
        let iterations = 0;
        const maxIterations = totalSlides + 5;
        
        while (pos < slideNumber && iterations < maxIterations) {
          slideShow.View.Next();
          await new Promise(resolve => setTimeout(resolve, 30));
          pos = queryCurrentSlide();
          iterations++;
        }
      }
    }
    
    await new Promise(resolve => setTimeout(resolve, 100));
    currentSlide = queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Now on slide ${currentSlide}`);
  },

  async getSlideInfo(): Promise<SlideInfo> {
    const actualSlide = queryCurrentSlide();
    
    if (actualSlide !== currentSlide) {
      currentSlide = actualSlide;
      currentAnimationStep = 0;
    }
    
    const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    
    // Find next visible slide for preview
    const nextVisible = presentationData 
      ? getNextVisibleSlide(currentSlide, presentationData)
      : (currentSlide < totalSlides ? currentSlide + 1 : null);
    
    // Get notes for current and next slide
    const currentNotes = slideData?.notes || '';
    const nextSlideData = nextVisible && presentationData 
      ? getSlideData(nextVisible, presentationData) 
      : null;
    const nextNotes = nextSlideData?.notes || '';
    
    // Check if we're on the last visible slide
    const visibleSlides = presentationData?.visibleSlides || [];
    const isLastSlide = visibleSlides.length > 0 
      ? currentSlide === visibleSlides[visibleSlides.length - 1]
      : currentSlide >= totalSlides;
    
    return {
      currentSlide,
      totalSlides,
      animationStep: currentAnimationStep,
      animationsOnSlide,
      nextVisibleSlide: nextVisible,
      isLastSlide,
      currentNotes,
      nextNotes,
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
      try {
        presentation.Close();
      } catch {
        // Ignore
      }
      presentation = null;
    }
    if (pptApp) {
      try {
        pptApp.Quit();
      } catch {
        // Ignore
      }
      pptApp = null;
    }
    
    // Clean up temp copy
    if (localPresentationCopy) {
      try {
        await unlink(localPresentationCopy);
      } catch {
        // Ignore
      }
      localPresentationCopy = '';
    }
    
    // Reset state
    slideShow = null;
    presentationData = null;
    currentSlide = 1;
    currentAnimationStep = 0;
  },
};
