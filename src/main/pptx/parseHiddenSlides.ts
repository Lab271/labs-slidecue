import { exec } from 'child_process';
import { promisify } from 'util';

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
    
    console.log('Parsing presentation.xml for hidden slides...');
    
    // Parse the XML to find hidden slides
    // Hidden slides have show="0" in their sldId entry
    // Format: <p:sldId id="256" r:id="rId2" show="0"/>
    
    // Find all sldId entries and check each one
    const sldIdRegex = /<p:sldId[^>]*>/g;
    const matches = stdout.match(sldIdRegex) || [];
    
    console.log(`Found ${matches.length} slide entries in presentation.xml`);
    
    matches.forEach((match, index) => {
      // Check if this slide has show="0" (hidden)
      if (match.includes('show="0"')) {
        // Slide indices are 1-based
        hiddenSlides.push(index + 1);
        console.log(`Slide ${index + 1} is hidden: ${match}`);
      }
    });
    
    // If no hidden slides found via show="0", check individual slide XML files
    // Some versions of PowerPoint store the hidden flag in the slide's own XML
    if (hiddenSlides.length === 0) {
      console.log('No hidden slides found via presentation.xml, checking individual slides...');
      
      const { stdout: fileList } = await execAsync(
        `unzip -l "${pptxPath}" 2>/dev/null | grep -E "ppt/slides/slide[0-9]+\\.xml" | awk '{print $4}'`
      );
      
      const slideFiles = fileList.trim().split('\n').filter(f => f);
      
      for (const slideFile of slideFiles) {
        const match = slideFile.match(/slide(\d+)\.xml$/);
        if (!match) continue;
        const slideNum = parseInt(match[1], 10);
        
        try {
          const { stdout: slideXml } = await execAsync(
            `unzip -p "${pptxPath}" "${slideFile}" 2>/dev/null`
          );
          
          // Check for show="0" in the slide itself
          if (slideXml.includes('show="0"') || slideXml.includes('show="false"')) {
            hiddenSlides.push(slideNum);
            console.log(`Slide ${slideNum} is hidden (found in slide XML)`);
          }
        } catch {
          // Ignore individual slide errors
        }
      }
    }
    
    console.log('Hidden slides detected:', hiddenSlides);
  } catch (e) {
    console.error('Failed to parse hidden slides:', e);
  }
  
  return hiddenSlides.sort((a, b) => a - b);
}

/**
 * Parse a PPTX file to count click-triggered animations per slide.
 * Returns a map of slideNumber -> number of click animations
 */
export async function getAnimationsPerSlide(pptxPath: string): Promise<Map<number, number>> {
  const animationsMap = new Map<number, number>();
  
  try {
    // List all slide XML files in the PPTX
    const { stdout: fileList } = await execAsync(
      `unzip -l "${pptxPath}" 2>/dev/null | grep "ppt/slides/slide[0-9]*.xml" | awk '{print $4}'`
    );
    
    const slideFiles = fileList.trim().split('\n').filter(f => f);
    console.log('Found slide files:', slideFiles);
    
    for (const slideFile of slideFiles) {
      // Extract slide number from filename (slide1.xml -> 1)
      const match = slideFile.match(/slide(\d+)\.xml$/);
      if (!match) continue;
      const slideNum = parseInt(match[1], 10);
      
      // Extract slide XML
      const { stdout: slideXml } = await execAsync(
        `unzip -p "${pptxPath}" "${slideFile}" 2>/dev/null`
      );
      
      // Count click-triggered animations
      // Look for <p:cTn> elements with nodeType="clickEffect" or
      // <p:seq> elements which are click sequences
      // Each <p:cTn clickToAdv="1"> or <p:seq> with concurrent="0" is a click-triggered animation
      
      // Count main sequence click triggers - these are in <p:seq> elements
      // Each <p:par> directly under <p:childTnLst> in <p:seq> is typically one click
      const seqMatches = slideXml.match(/<p:seq[^>]*>[\s\S]*?<\/p:seq>/g) || [];
      
      let clickCount = 0;
      for (const seq of seqMatches) {
        // Count <p:par> elements at the top level of the sequence's childTnLst
        // These represent click-triggered animation groups
        const parMatches = seq.match(/<p:cTn[^>]*nodeType="clickEffect"/g) || [];
        clickCount += parMatches.length;
      }
      
      // Alternative: count all nodeType="clickEffect" as a simpler approach
      const allClickEffects = slideXml.match(/nodeType="clickEffect"/g) || [];
      clickCount = allClickEffects.length;
      
      if (clickCount > 0) {
        animationsMap.set(slideNum, clickCount);
        console.log(`Slide ${slideNum}: ${clickCount} click animation(s)`);
      }
    }
  } catch (e) {
    console.error('Failed to parse animations:', e);
  }
  
  return animationsMap;
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
