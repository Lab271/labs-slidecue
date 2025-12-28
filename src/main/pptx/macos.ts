import { exec } from 'child_process';
import { promisify } from 'util';
import { readdir, access, mkdir, unlink, copyFile } from 'fs/promises';
import { join, basename } from 'path';
import { tmpdir } from 'os';
import { PowerPointAutomation, SlideInfo, SlideMetadata, ProgressCallback } from './types';
import { parsePresentationData, PresentationData, getNextVisibleSlide, getSlideData } from './slideParser';

const execAsync = promisify(exec);

// Presentation state
let presentationData: PresentationData | null = null;
let currentSlide = 1;
let currentAnimationStep = 0;
let totalSlides = 1;
let currentPresentationPath = '';
let currentThumbsDir = '';
let localPresentationCopy = '';

function runAppleScript(script: string): Promise<string> {
  return execAsync(`osascript <<'EOF'
${script}
EOF`).then(({ stdout }) => stdout.trim());
}

function sendKeyCode(keyCode: number): Promise<string> {
  return runAppleScript(`
tell application "Microsoft PowerPoint" to activate
delay 0.1
tell application "System Events"
  key code ${keyCode}
end tell
  `);
}

/**
 * Query PowerPoint for the actual current slide number
 */
async function queryCurrentSlide(): Promise<number> {
  try {
    const result = await runAppleScript(`
tell application "Microsoft PowerPoint"
  try
    set ss to slide show window 1
    set currentSlideIndex to slide index of slide of slide show view of ss
    return currentSlideIndex
  on error
    return 1
  end try
end tell
    `);
    return parseInt(result.trim(), 10) || 1;
  } catch {
    return currentSlide;
  }
}

export const macOSAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      await access('/Applications/Microsoft PowerPoint.app');
      return true;
    } catch {
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
    
    // Parse presentation data (slides, hidden, animations)
    console.log('Parsing presentation data...');
    presentationData = await parsePresentationData(localPresentationCopy);
    
    console.log('Presentation data:', JSON.stringify(presentationData, null, 2));
    
    // Use 'open' command which handles permissions better than AppleScript
    console.log('Opening presentation...');
    await execAsync(`open -a "Microsoft PowerPoint" "${localPresentationCopy}"`);
    
    // Wait for file to open
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Get total slides from PowerPoint to verify
    try {
      const total = await runAppleScript(`
tell application "Microsoft PowerPoint"
  return count of slides of active presentation
end tell
      `);
      totalSlides = parseInt(total, 10) || presentationData.totalSlides;
    } catch {
      totalSlides = presentationData.totalSlides;
    }
    
    currentSlide = 1;
    currentAnimationStep = 0;
    
    console.log(`Opened presentation with ${totalSlides} slides`);
    console.log(`Visible slides: ${presentationData.visibleSlides.join(', ')}`);
    console.log(`Hidden slides: ${presentationData.hiddenSlides.join(', ')}`);
  },

  async exportThumbnails(outputDir: string, onProgress?: ProgressCallback) {
    currentThumbsDir = outputDir;
    await mkdir(outputDir, { recursive: true });
    
    // Use the local copy path for conversion
    const pptxPath = localPresentationCopy || currentPresentationPath;
    const pdfPath = join(outputDir, 'presentation.pdf');
    
    console.log('Converting presentation to PDF and images...');
    onProgress?.(1, totalSlides);
    
    try {
      // Use LibreOffice headless conversion (most reliable)
      console.log('Using LibreOffice to convert PPTX to PDF...');
      await execAsync(`/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir "${outputDir}" "${pptxPath}"`);
      
      // Rename the output file (LibreOffice uses the original filename)
      const baseName = basename(pptxPath).replace(/\.(pptx?|PPTX?)$/, '');
      const generatedPdf = join(outputDir, baseName + '.pdf');
      
      try {
        await access(generatedPdf);
        if (generatedPdf !== pdfPath) {
          await execAsync(`mv "${generatedPdf}" "${pdfPath}"`);
        }
      } catch {
        // PDF might already be named correctly
      }
      
      // Check if PDF was created
      await access(pdfPath);
      console.log('PDF created, converting to PNG...');
      
      // Convert PDF pages to PNG using pdftoppm
      await execAsync(`pdftoppm -png -r 150 "${pdfPath}" "${outputDir}/slide"`);
      
      // pdftoppm creates slide-1.png, slide-2.png, etc.
      // LibreOffice only exports visible slides, so we need to rename them
      // to match the actual PowerPoint slide numbers
      const files = await readdir(outputDir);
      const pngFiles = files.filter(f => f.match(/^slide-\d+\.png$/)).sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)![0]);
        const numB = parseInt(b.match(/\d+/)![0]);
        return numA - numB;
      });
      
      // Use visible slides from parsed data
      const visibleSlides = presentationData?.visibleSlides || [];
      
      console.log('Visible slide numbers:', visibleSlides);
      console.log('PNG files to rename:', pngFiles);
      
      // Rename each exported slide to its actual PowerPoint slide number
      for (let i = 0; i < pngFiles.length; i++) {
        const file = pngFiles[i];
        const actualSlideNum = visibleSlides[i] || (i + 1);
        const newName = `slide_${String(actualSlideNum).padStart(3, '0')}.png`;
        await execAsync(`mv "${join(outputDir, file)}" "${join(outputDir, newName)}"`);
        console.log(`Renamed ${file} -> ${newName}`);
      }
      
      // Clean up PDF
      await unlink(pdfPath).catch(() => {});
      
    } catch (e) {
      console.error('Export failed:', e);
    }
    
    // Find all generated PNG files
    let thumbnails: string[] = [];
    try {
      const files = await readdir(outputDir);
      const pngs = files
        .filter(f => f.endsWith('.png'))
        .sort();
      
      for (const png of pngs) {
        thumbnails.push(join(outputDir, png));
      }
    } catch {
      // Ignore
    }
    
    console.log('Generated thumbnails:', thumbnails);
    
    // Return metadata
    const result: SlideMetadata = {
      thumbnails,
      totalSlides,
      hiddenSlides: presentationData?.hiddenSlides || [],
      visibleSlides: presentationData?.visibleSlides || [],
    };
    
    return result;
  },

  async startSlideshow() {
    currentSlide = 1;
    currentAnimationStep = 0;
    
    await runAppleScript(`
tell application "Microsoft PowerPoint"
  activate
end tell
delay 0.2
tell application "System Events"
  keystroke return using {command down, shift down}
end tell
    `);
    await new Promise(resolve => setTimeout(resolve, 500));
    
    // PowerPoint may skip hidden slide 1, query actual position
    currentSlide = await queryCurrentSlide();
  },

  async nextSlide() {
    const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    
    // Send the keystroke
    await sendKeyCode(124); // Right arrow
    
    // Wait for PowerPoint to process
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Check if we advanced an animation or moved to next slide
    if (currentAnimationStep < animationsOnSlide) {
      // Might be an animation click
      currentAnimationStep++;
    }
    
    // Query PowerPoint for actual slide number
    const actualSlide = await queryCurrentSlide();
    
    if (actualSlide !== currentSlide) {
      // We moved to a new slide
      currentSlide = actualSlide;
      currentAnimationStep = 0;
      console.log(`Moved to slide ${currentSlide}`);
    } else {
      console.log(`Animation ${currentAnimationStep}/${animationsOnSlide} on slide ${currentSlide}`);
    }
  },

  async prevSlide() {
    await sendKeyCode(123); // Left arrow
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // Query PowerPoint for actual slide number
    currentSlide = await queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Moved to slide ${currentSlide}`);
  },

  async gotoSlide(slideNumber: number) {
    console.log(`Going to slide ${slideNumber}...`);
    
    // PowerPoint's AppleScript "go to slide" command with a number parameter
    // doesn't work (Dutch localization issue?), but go to first/last/next/previous work.
    // We use a single AppleScript with a repeat loop for efficiency.
    try {
      if (slideNumber === 1) {
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to first slide
  end tell
end tell
        `);
      } else if (slideNumber >= totalSlides) {
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to last slide
  end tell
end tell
        `);
      } else {
        // Go to first slide, then advance until we reach target position
        // The repeat loop runs in a single AppleScript call for speed
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to first slide
    repeat until (current show position) ≥ ${slideNumber}
      go to next slide
    end repeat
  end tell
end tell
        `);
      }
    } catch (e) {
      console.error('Error in gotoSlide:', e);
    }
    
    await new Promise(resolve => setTimeout(resolve, 100));
    currentSlide = await queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Now on slide ${currentSlide}`);
  },

  async getSlideInfo(): Promise<SlideInfo> {
    // Query PowerPoint for actual slide
    const actualSlide = await queryCurrentSlide();
    
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
    await sendKeyCode(53); // Escape key
  },

  async closePresentation() {
    await runAppleScript(`
tell application "Microsoft PowerPoint"
  close active presentation saving no
end tell
    `);
    
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
    presentationData = null;
    currentSlide = 1;
    currentAnimationStep = 0;
  },
};
