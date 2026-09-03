// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { readdir, mkdir, unlink, copyFile } from 'fs/promises';
import { join, basename } from 'path';
import { tmpdir } from 'os';
import log from 'electron-log';
import { PowerPointAutomation, SlideInfo, SlideMetadata, ProgressCallback } from './types';
import { serializeAutomation } from './serialize';
import { parsePresentationData, PresentationData, getNextVisibleSlide, getSlideData } from './slideParser';

// Try to load winax
let winax: any = null;
try {
  winax = require('winax');
  log.info('[Windows-WinAX] Loaded winax module successfully');
} catch (error) {
  log.error('[Windows-WinAX] Failed to load winax:', error);
}

// Presentation state
let pptApp: any = null;
let presentation: any = null;
let slideShow: any = null;
let presentationData: PresentationData | null = null;
let currentSlide = 1;
let currentAnimationStep = 0;
let totalSlides = 1;
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

const winaxBackend: PowerPointAutomation = {
  async checkInstalled() {
    if (!winax) {
      log.error('[Windows-WinAX] winax module not available');
      return false;
    }

    try {
      log.info('[Windows-WinAX] Checking if PowerPoint is installed');
      const testApp = new winax.Object('PowerPoint.Application');
      const version = testApp.Version;
      testApp.Quit();
      log.info('[Windows-WinAX] PowerPoint is installed, version:', version);
      return true;
    } catch (error) {
      log.error('[Windows-WinAX] PowerPoint check failed:', error);
      return false;
    }
  },

  async openPresentation(filePath: string) {
    if (!winax) {
      throw new Error('winax module not available');
    }

    try {
      log.info('[Windows-WinAX] Opening presentation:', filePath);
      
      // Copy file to local temp directory
      const tempDir = join(tmpdir(), 'slidecue-presentations');
      await mkdir(tempDir, { recursive: true });
      localPresentationCopy = join(tempDir, basename(filePath));
      
      log.info('[Windows-WinAX] Copying presentation to temp location:', localPresentationCopy);
      await copyFile(filePath, localPresentationCopy);
      
      // Parse presentation data
      log.info('[Windows-WinAX] Parsing presentation data');
      presentationData = await parsePresentationData(localPresentationCopy);
      
      log.info('[Windows-WinAX] Presentation data:', {
        totalSlides: presentationData.totalSlides,
        visibleSlides: presentationData.visibleSlides,
        hiddenSlides: presentationData.hiddenSlides
      });
      
      // Open in PowerPoint via COM
      log.info('[Windows-WinAX] Opening presentation in PowerPoint');
      pptApp = new winax.Object('PowerPoint.Application');
      pptApp.Visible = 1; // msoTrue
      presentation = pptApp.Presentations.Open(localPresentationCopy);
      
      // Get total slides
      totalSlides = presentation.Slides.Count || presentationData.totalSlides;
      currentSlide = 1;
      currentAnimationStep = 0;
      
      log.info('[Windows-WinAX] Opened presentation with', totalSlides, 'slides');
      log.info('[Windows-WinAX] Visible slides:', presentationData.visibleSlides.join(', '));
      log.info('[Windows-WinAX] Hidden slides:', presentationData.hiddenSlides.join(', '));
    } catch (error) {
      log.error('[Windows-WinAX] Failed to open presentation:', error);
      throw error;
    }
  },

  async exportThumbnails(outputDir: string, onProgress?: ProgressCallback): Promise<SlideMetadata> {
    try {
      log.info('[Windows-WinAX] Exporting thumbnails to:', outputDir);
      await mkdir(outputDir, { recursive: true });
      
      const visibleSlides = presentationData?.visibleSlides || [];
      onProgress?.(1, visibleSlides.length);
      
      // Export only visible slides
      for (let i = 0; i < visibleSlides.length; i++) {
        const slideNum = visibleSlides[i];
        const slide = presentation.Slides.Item(slideNum);
        const filePath = join(outputDir, `slide_${String(slideNum).padStart(3, '0')}.png`);
        
        try {
          slide.Export(filePath, 'PNG', 1920, 1080);
          log.info('[Windows-WinAX] Exported slide', slideNum);
        } catch (e) {
          log.error('[Windows-WinAX] Failed to export slide', slideNum, ':', e);
        }
        
        onProgress?.(i + 1, visibleSlides.length);
      }
      
      // Get paths to exported files
      const files = await readdir(outputDir);
      const paths = files
        .filter(f => f.endsWith('.png'))
        .sort()
        .map(f => join(outputDir, f));
      
      log.info('[Windows-WinAX] Generated', paths.length, 'thumbnail files');
      
      return {
        thumbnails: paths,
        totalSlides,
        hiddenSlides: presentationData?.hiddenSlides || [],
        visibleSlides: presentationData?.visibleSlides || [],
      };
    } catch (error) {
      log.error('[Windows-WinAX] Failed to export thumbnails:', error);
      throw error;
    }
  },

  async startSlideshow() {
    try {
      log.info('[Windows-WinAX] Starting slideshow');
      currentSlide = 1;
      currentAnimationStep = 0;
      
      const settings = presentation.SlideShowSettings;
      settings.StartingSlide = 1;
      settings.EndingSlide = totalSlides;
      slideShow = settings.Run();
      
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // PowerPoint may skip hidden slide 1, query actual position
      currentSlide = queryCurrentSlide();
      log.info('[Windows-WinAX] Slideshow started at slide', currentSlide);
    } catch (error) {
      log.error('[Windows-WinAX] Failed to start slideshow:', error);
      throw error;
    }
  },

  async nextSlide() {
    try {
      const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
      const animationsOnSlide = slideData?.animationClicks || 0;
      
      log.info('[Windows-WinAX] Next slide (current:', currentSlide, 'animation:', currentAnimationStep, '/', animationsOnSlide, ')');
      
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
        log.info('[Windows-WinAX] Moved to slide', currentSlide);
      } else {
        log.info('[Windows-WinAX] Animation', currentAnimationStep, '/', animationsOnSlide, 'on slide', currentSlide);
      }
    } catch (error) {
      log.error('[Windows-WinAX] Failed to advance slide:', error);
    }
  },

  async prevSlide() {
    try {
      log.info('[Windows-WinAX] Previous slide (current:', currentSlide, ')');
      
      if (slideShow?.View) {
        slideShow.View.Previous();
      }
      
      await new Promise(resolve => setTimeout(resolve, 100));
      
      // Logging the move re-reads the state after the await, so the assignment
      // below cannot be based on the value captured before the COM call.
      const actualSlide = queryCurrentSlide();
      log.info('[Windows-WinAX] Moved from slide', currentSlide, 'to slide', actualSlide);
      currentSlide = actualSlide;
      currentAnimationStep = 0;
    } catch (error) {
      log.error('[Windows-WinAX] Failed to go to previous slide:', error);
    }
  },

  async gotoSlide(slideNumber: number) {
    try {
      log.info('[Windows-WinAX] Going to slide', slideNumber);
      
      if (slideShow?.View) {
        try {
          slideShow.View.GotoSlide(slideNumber);
        } catch (e) {
          log.error('[Windows-WinAX] GotoSlide failed, trying navigation approach:', e);
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
      log.info('[Windows-WinAX] Now on slide', currentSlide);
    } catch (error) {
      log.error('[Windows-WinAX] Failed to go to slide:', error);
    }
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
    try {
      log.info('[Windows-WinAX] Stopping slideshow');
      if (slideShow?.View) {
        slideShow.View.Exit();
        slideShow = null;
      }
    } catch (error) {
      log.error('[Windows-WinAX] Failed to stop slideshow:', error);
    }
  },

  async closePresentation() {
    try {
      log.info('[Windows-WinAX] Closing presentation');
      
      if (presentation) {
        try {
          presentation.Close();
        } catch (error) {
          log.error('[Windows-WinAX] Error closing presentation:', error);
        }
        presentation = null;
      }
      
      if (pptApp) {
        try {
          pptApp.Quit();
        } catch (error) {
          log.error('[Windows-WinAX] Error quitting PowerPoint:', error);
        }
        pptApp = null;
      }
      
      // Clean up temp copy. The path is cleared before the unlink is awaited,
      // so the field is never left pointing at a file that is on its way out.
      const tempCopy = localPresentationCopy;
      if (tempCopy) {
        localPresentationCopy = '';
        try {
          await unlink(tempCopy);
          log.info('[Windows-WinAX] Deleted temp presentation copy');
        } catch (error) {
          log.error('[Windows-WinAX] Failed to delete temp copy:', error);
        }
      }
      
      // Reset state
      slideShow = null;
      presentationData = null;
      currentSlide = 1;
      currentAnimationStep = 0;
      totalSlides = 1;
    } catch (error) {
      log.error('[Windows-WinAX] Failed to close presentation:', error);
    }
  },
};

// One command at a time: the module-level state above is only consistent if
// nothing interleaves with it. See serialize.ts.
export const windowsWinaxAutomation = serializeAutomation(winaxBackend);
