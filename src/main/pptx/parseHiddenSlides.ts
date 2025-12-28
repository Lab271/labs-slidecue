import { exec } from 'child_process';
import { promisify } from 'util';
import { readFile } from 'fs/promises';

const execAsync = promisify(exec);

/**
 * Parse a PPTX file to find which slides are hidden.
 * PPTX is a ZIP file - we extract presentation.xml and check for show="0" attributes.
 */
export async function getHiddenSlides(pptxPath: string): Promise<number[]> {
  const hiddenSlides: number[] = [];
  
  try {
    // Extract presentation.xml from the PPTX (which is a ZIP file)
    const { stdout } = await execAsync(
      `unzip -p "${pptxPath}" "ppt/presentation.xml" 2>/dev/null`
    );
    
    // Parse the XML to find hidden slides
    // Hidden slides have show="0" in their sldId entry
    // Format: <p:sldId id="256" r:id="rId2" show="0"/>
    
    // Find all sldId entries
    const sldIdRegex = /<p:sldId[^>]*>/g;
    const matches = stdout.match(sldIdRegex) || [];
    
    matches.forEach((match, index) => {
      // Check if this slide has show="0" (hidden)
      if (match.includes('show="0"')) {
        // Slide indices are 1-based
        hiddenSlides.push(index + 1);
      }
    });
    
    console.log('Hidden slides detected:', hiddenSlides);
  } catch (e) {
    console.error('Failed to parse hidden slides:', e);
  }
  
  return hiddenSlides;
}

/**
 * Convert a "visible slide index" (counting only non-hidden slides) 
 * to an "actual slide index" (counting all slides).
 */
export function visibleToActualIndex(visibleIndex: number, hiddenSlides: number[]): number {
  let actualIndex = 0;
  let visibleCount = 0;
  
  while (visibleCount < visibleIndex) {
    actualIndex++;
    if (!hiddenSlides.includes(actualIndex)) {
      visibleCount++;
    }
  }
  
  return actualIndex;
}

/**
 * Convert an "actual slide index" to a "visible slide index".
 */
export function actualToVisibleIndex(actualIndex: number, hiddenSlides: number[]): number {
  let visibleIndex = 0;
  
  for (let i = 1; i <= actualIndex; i++) {
    if (!hiddenSlides.includes(i)) {
      visibleIndex++;
    }
  }
  
  return visibleIndex;
}
