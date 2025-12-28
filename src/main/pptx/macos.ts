import { exec } from 'child_process';
import { promisify } from 'util';
import { readdir, access, mkdir, unlink, copyFile } from 'fs/promises';
import { join, basename } from 'path';
import { tmpdir } from 'os';
import { PowerPointAutomation, SlideInfo, SlideMetadata, ProgressCallback } from './types';
import { getHiddenSlides, actualToVisibleIndex, visibleToActualIndex } from './parseHiddenSlides';

const execAsync = promisify(exec);

// Track slide state since we can't reliably query PowerPoint
let currentVisibleSlide = 1; // Current slide in visible-only count
let totalSlides = 1;
let hiddenSlides: number[] = [];
let visibleSlideCount = 1;
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
    
    // Detect hidden slides before opening
    hiddenSlides = await getHiddenSlides(localPresentationCopy);
    
    // Use 'open' command which handles permissions better than AppleScript
    console.log('Opening presentation...');
    await execAsync(`open -a "Microsoft PowerPoint" "${localPresentationCopy}"`);
    
    // Wait for file to open
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    // Get total slides
    try {
      const total = await runAppleScript(`
tell application "Microsoft PowerPoint"
  return count of slides of active presentation
end tell
      `);
      totalSlides = parseInt(total, 10) || 1;
    } catch {
      totalSlides = 1;
    }
    
    // Calculate visible slide count
    visibleSlideCount = totalSlides - hiddenSlides.length;
    currentVisibleSlide = 1;
    
    console.log(`Opened presentation with ${totalSlides} slides (${hiddenSlides.length} hidden, ${visibleSlideCount} visible)`);
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
      
      // pdftoppm creates slide-1.png, slide-2.png, etc. - rename to slide_001.png format
      const files = await readdir(outputDir);
      for (const file of files) {
        const match = file.match(/^slide-(\d+)\.png$/);
        if (match) {
          const num = parseInt(match[1], 10);
          const newName = `slide_${String(num).padStart(3, '0')}.png`;
          await execAsync(`mv "${join(outputDir, file)}" "${join(outputDir, newName)}"`);
        }
      }
      
      // Clean up PDF
      await unlink(pdfPath).catch(() => {});
      
    } catch (e) {
      console.error('Export failed:', e);
    }
    
    // Find all generated PNG files
    let allThumbnails: string[] = [];
    try {
      const files = await readdir(outputDir);
      const pngs = files
        .filter(f => f.endsWith('.png'))
        .sort();
      
      for (const png of pngs) {
        allThumbnails.push(join(outputDir, png));
      }
    } catch {
      // Ignore
    }
    
    console.log('Generated thumbnails (all):', allThumbnails);
    
    // Filter out hidden slide thumbnails for the visible-only list
    const visibleThumbnails = allThumbnails.filter((_, index) => {
      const slideNum = index + 1; // 1-based slide number
      return !hiddenSlides.includes(slideNum);
    });
    
    console.log('Visible thumbnails:', visibleThumbnails);
    console.log('Hidden slides:', hiddenSlides);
    
    // Update module-level counts based on actual generated thumbnails
    // This ensures getSlideInfo() returns accurate counts
    totalSlides = allThumbnails.length;
    visibleSlideCount = visibleThumbnails.length;
    
    console.log(`Updated slide counts: ${visibleSlideCount} visible of ${totalSlides} total`);
    
    // Return metadata - only include visible thumbnails
    const result: SlideMetadata = {
      thumbnails: visibleThumbnails,
      totalSlides,
      hiddenSlides,
      visibleSlideCount,
    };
    
    return result;
  },

  async startSlideshow() {
    currentVisibleSlide = 1;
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
  },

  async nextSlide() {
    await sendKeyCode(124); // Right arrow
    // PowerPoint automatically skips hidden slides, so we just track visible count
    if (currentVisibleSlide < visibleSlideCount) {
      currentVisibleSlide++;
    }
  },

  async prevSlide() {
    await sendKeyCode(123); // Left arrow
    // PowerPoint automatically skips hidden slides going backward too
    if (currentVisibleSlide > 1) {
      currentVisibleSlide--;
    }
  },

  async gotoSlide(visibleIndex: number) {
    // Convert visible index to actual slide number for PowerPoint
    const actualIndex = visibleToActualIndex(visibleIndex, hiddenSlides);
    await runAppleScript(`
tell application "Microsoft PowerPoint" to activate
delay 0.1
tell application "System Events"
  keystroke "${actualIndex}"
  delay 0.1
  keystroke return
end tell
    `);
    currentVisibleSlide = visibleIndex;
  },

  async getSlideInfo(): Promise<SlideInfo> {
    return {
      current: currentVisibleSlide,
      total: visibleSlideCount,
      hiddenSlides: hiddenSlides,
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
  },
};
