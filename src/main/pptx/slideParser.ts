import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

export interface SlideData {
  slideNumber: number;      // 1-based actual slide number
  name: string;             // Slide name/title if available
  hidden: boolean;          // Is this slide hidden?
  animationClicks: number;  // Number of click-triggered animations
  notes: string;            // Speaker notes
}

export interface PresentationData {
  slides: SlideData[];
  totalSlides: number;
  visibleSlides: number[];  // Array of visible slide numbers
  hiddenSlides: number[];   // Array of hidden slide numbers
}

/**
 * Parse a PPTX file to extract all slide metadata
 */
export async function parsePresentationData(pptxPath: string): Promise<PresentationData> {
  const slides: SlideData[] = [];
  const visibleSlides: number[] = [];
  const hiddenSlides: number[] = [];
  
  try {
    // Get list of slide files
    const { stdout: fileList } = await execAsync(
      `unzip -l "${pptxPath}" 2>/dev/null | grep -E "ppt/slides/slide[0-9]+\\.xml" | awk '{print $4}' | sort -V`
    );
    
    const slideFiles = fileList.trim().split('\n').filter(f => f);
    console.log(`Found ${slideFiles.length} slide files`);
    
    // Extract presentation.xml to check for hidden slides
    let hiddenInPresentation: number[] = [];
    try {
      const { stdout: presXml } = await execAsync(
        `unzip -p "${pptxPath}" "ppt/presentation.xml" 2>/dev/null`
      );
      
      // Find sldId entries with show="0"
      const sldIdRegex = /<p:sldId[^>]*>/g;
      const matches = presXml.match(sldIdRegex) || [];
      matches.forEach((match, index) => {
        if (match.includes('show="0"')) {
          hiddenInPresentation.push(index + 1);
        }
      });
    } catch (e) {
      console.log('Could not parse presentation.xml for hidden slides');
    }
    
    // Parse each slide
    for (const slideFile of slideFiles) {
      const match = slideFile.match(/slide(\d+)\.xml$/);
      if (!match) continue;
      
      const slideNum = parseInt(match[1], 10);
      
      try {
        const { stdout: slideXml } = await execAsync(
          `unzip -p "${pptxPath}" "${slideFile}" 2>/dev/null`
        );
        
        // Check if hidden (in slide XML itself)
        let isHidden = hiddenInPresentation.includes(slideNum);
        if (slideXml.includes('show="0"') || slideXml.includes('show="false"')) {
          isHidden = true;
        }
        
        // Count click-triggered animations
        // Look for nodeType="clickEffect" in timing nodes
        const clickEffects = (slideXml.match(/nodeType="clickEffect"/g) || []).length;
        
        // Try to get slide name/title
        let name = `Slide ${slideNum}`;
        const titleMatch = slideXml.match(/<a:t>([^<]+)<\/a:t>/);
        if (titleMatch && titleMatch[1].length < 100) {
          name = titleMatch[1].substring(0, 50);
        }
        
        // Try to get notes by looking up the relationship file first
        let notes = '';
        try {
          // Get the slide's relationship file to find the actual notes file
          const slideRelsFile = slideFile.replace('slides/', 'slides/_rels/').replace('.xml', '.xml.rels');
          const { stdout: relsXml } = await execAsync(
            `unzip -p "${pptxPath}" "${slideRelsFile}" 2>/dev/null`
          );
          
          // Find the notesSlide relationship
          const notesMatch = relsXml.match(/Type="[^"]*notesSlide"[^>]*Target="([^"]+)"/);
          
          if (notesMatch && notesMatch[1]) {
            // Target is relative like "../notesSlides/notesSlide2.xml"
            const notesTarget = notesMatch[1];
            const notesFile = 'ppt/' + notesTarget.replace('../', '');
            
            console.log(`Slide ${slideNum} notes file: ${notesFile}`);
            
            const { stdout: notesXml } = await execAsync(
              `unzip -p "${pptxPath}" "${notesFile}" 2>/dev/null`
            );
            
            // The notes are in the shape with <p:ph type="body"/>
            const bodyMatch = notesXml.match(/<p:sp>.*?<p:ph type="body"[^>]*\/>.*?<p:txBody>(.*?)<\/p:txBody>.*?<\/p:sp>/s);
            
            if (bodyMatch && bodyMatch[1]) {
              const txBody = bodyMatch[1];
              // Extract all <a:t> text content
              const textMatches = txBody.match(/<a:t>([^<]*)<\/a:t>/g) || [];
              const textParts: string[] = [];
              
              for (const match of textMatches) {
                const text = match.replace(/<\/?a:t>/g, '');
                textParts.push(text);
              }
              
              notes = textParts.join('').trim();
            }
            
            // Limit to reasonable length
            if (notes.length > 1000) {
              notes = notes.substring(0, 1000) + '...';
            }
            
            console.log(`Slide ${slideNum} notes: "${notes}"`);
          } else {
            console.log(`Slide ${slideNum}: no notes relationship found`);
          }
        } catch (e) {
          // No notes for this slide
          console.log(`Slide ${slideNum}: no notes file`);
        }
        
        const slideData: SlideData = {
          slideNumber: slideNum,
          name,
          hidden: isHidden,
          animationClicks: clickEffects,
          notes,
        };
        
        slides.push(slideData);
        
        if (isHidden) {
          hiddenSlides.push(slideNum);
        } else {
          visibleSlides.push(slideNum);
        }
        
        console.log(`Slide ${slideNum}: hidden=${isHidden}, animations=${clickEffects}`);
        
      } catch (e) {
        console.error(`Error parsing slide ${slideNum}:`, e);
        // Add placeholder
        slides.push({
          slideNumber: slideNum,
          name: `Slide ${slideNum}`,
          hidden: false,
          animationClicks: 0,
          notes: '',
        });
        visibleSlides.push(slideNum);
      }
    }
    
  } catch (e) {
    console.error('Failed to parse presentation:', e);
  }
  
  // Sort by slide number
  slides.sort((a, b) => a.slideNumber - b.slideNumber);
  visibleSlides.sort((a, b) => a - b);
  hiddenSlides.sort((a, b) => a - b);
  
  return {
    slides,
    totalSlides: slides.length,
    visibleSlides,
    hiddenSlides,
  };
}

/**
 * Find the next visible slide after the given slide number
 */
export function getNextVisibleSlide(currentSlide: number, presentationData: PresentationData): number | null {
  const { visibleSlides, totalSlides } = presentationData;
  
  for (const slideNum of visibleSlides) {
    if (slideNum > currentSlide) {
      return slideNum;
    }
  }
  
  return null; // No more visible slides
}

/**
 * Find the previous visible slide before the given slide number
 */
export function getPrevVisibleSlide(currentSlide: number, presentationData: PresentationData): number | null {
  const { visibleSlides } = presentationData;
  
  for (let i = visibleSlides.length - 1; i >= 0; i--) {
    if (visibleSlides[i] < currentSlide) {
      return visibleSlides[i];
    }
  }
  
  return null; // No previous visible slides
}

/**
 * Get slide data by slide number
 */
export function getSlideData(slideNumber: number, presentationData: PresentationData): SlideData | null {
  return presentationData.slides.find(s => s.slideNumber === slideNumber) || null;
}
