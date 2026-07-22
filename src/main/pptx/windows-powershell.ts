// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { spawn, ChildProcess } from 'child_process';
import { readdir, mkdir, unlink, copyFile } from 'fs/promises';
import { join, basename } from 'path';
import { tmpdir } from 'os';
import path from 'path';
import log from 'electron-log';
import { app } from 'electron';
import { PowerPointAutomation, SlideInfo, SlideMetadata, ProgressCallback } from './types';
import { parsePresentationData, PresentationData, getNextVisibleSlide, getSlideData } from './slideParser';

interface PSCommand {
  action: string;
  filePath?: string;
  outputDir?: string;
  slideNumber?: number;
}

interface PSResponse {
  status: string;
  data?: string;
  error?: string;
}

// PowerShell bridge process manager
class PowerShellBridge {
  private process: ChildProcess | null = null;
  private isReady = false;
  private commandQueue: Array<{ 
    resolve: (value: PSResponse) => void; 
    reject: (error: Error) => void;
    timeout: NodeJS.Timeout;
  }> = [];
  private currentCommand: { 
    resolve: (value: PSResponse) => void; 
    reject: (error: Error) => void;
    timeout: NodeJS.Timeout;
  } | null = null;

  async start(): Promise<void> {
    if (this.process) {
      return;
    }

    log.info('[Windows] Starting PowerShell bridge');
    
    const fs = require('fs');
    
    // Try multiple possible locations for the PowerShell script
    const possiblePaths = [
      // Packaged location
      app.isPackaged ? path.join(process.resourcesPath, 'powerpoint-bridge.ps1') : null,
      // Build output location
      path.join(__dirname, 'powerpoint-bridge.ps1'),
      // Source location (fallback for development)
      path.join(__dirname, '..', '..', 'src', 'main', 'pptx', 'powerpoint-bridge.ps1'),
      // Alternative source location
      path.join(process.cwd(), 'src', 'main', 'pptx', 'powerpoint-bridge.ps1'),
    ].filter(Boolean);

    log.info('[Windows] __dirname:', __dirname);
    log.info('[Windows] app.isPackaged:', app.isPackaged);
    log.info('[Windows] process.cwd():', process.cwd());
    
    let scriptPath: string | null = null;
    for (const testPath of possiblePaths) {
      log.info('[Windows] Checking path:', testPath);
      if (fs.existsSync(testPath)) {
        scriptPath = testPath;
        log.info('[Windows] Found PowerShell script at:', scriptPath);
        break;
      }
    }
    
    if (!scriptPath) {
      log.error('[Windows] PowerShell script not found in any of these locations:', possiblePaths);
      throw new Error('PowerShell script not found');
    }

    this.process = spawn('powershell', [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath
    ], {
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true
    });

    this.process.stdout?.setEncoding('utf8');
    this.process.stderr?.setEncoding('utf8');

    // Handle stdout
    let buffer = '';
    this.process.stdout?.on('data', (data: string) => {
      buffer += data;
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;

        try {
          const response: PSResponse = JSON.parse(trimmed);
          log.info('[Windows] PowerShell response:', response);

          if (response.status === 'ready') {
            this.isReady = true;
            log.info('[Windows] PowerShell bridge is ready');
          } else if (this.currentCommand) {
            clearTimeout(this.currentCommand.timeout);
            this.currentCommand.resolve(response);
            this.currentCommand = null;
            this.processQueue();
          }
        } catch (error) {
          log.error('[Windows] Failed to parse PowerShell response:', trimmed, error);
        }
      }
    });

    this.process.stderr?.on('data', (data: string) => {
      log.error('[Windows] PowerShell stderr:', data);
    });

    this.process.on('error', (error) => {
      log.error('[Windows] PowerShell process error:', error);
      if (this.currentCommand) {
        clearTimeout(this.currentCommand.timeout);
        this.currentCommand.reject(error);
        this.currentCommand = null;
      }
    });

    this.process.on('exit', (code) => {
      log.info('[Windows] PowerShell bridge exited with code:', code);
      this.process = null;
      this.isReady = false;
    });

    // Wait for ready signal (with timeout)
    await new Promise<void>((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error('PowerShell bridge startup timeout'));
      }, 10000);

      const checkReady = setInterval(() => {
        if (this.isReady) {
          clearInterval(checkReady);
          clearTimeout(timeout);
          resolve();
        }
      }, 100);
    });
  }

  private processQueue(): void {
    if (this.currentCommand || this.commandQueue.length === 0) {
      return;
    }

    this.currentCommand = this.commandQueue.shift()!;
  }

  async sendCommand(command: PSCommand, timeoutMs = 30000): Promise<PSResponse> {
    if (!this.process || !this.isReady) {
      await this.start();
    }

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        log.error('[Windows] Command timeout:', command);
        reject(new Error(`Command timeout: ${command.action}`));
        if (this.currentCommand) {
          this.currentCommand = null;
          this.processQueue();
        }
      }, timeoutMs);

      this.commandQueue.push({ resolve, reject, timeout });
      
      const commandStr = JSON.stringify(command) + '\n';
      log.info('[Windows] Sending command:', commandStr.trim());
      this.process!.stdin?.write(commandStr);

      if (!this.currentCommand) {
        this.processQueue();
      }
    });
  }

  async stop(): Promise<void> {
    if (this.process) {
      log.info('[Windows] Stopping PowerShell bridge');
      try {
        await this.sendCommand({ action: 'quit' }, 5000);
      } catch (error) {
        log.error('[Windows] Error sending quit command:', error);
      }
      this.process.kill();
      this.process = null;
      this.isReady = false;
    }
  }
}

const bridge = new PowerShellBridge();

// Presentation state
let presentationData: PresentationData | null = null;
let currentSlide = 1;
let currentAnimationStep = 0;
let totalSlides = 1;
let currentPresentationPath = '';
let localPresentationCopy = '';

export const windowsAutomation: PowerPointAutomation = {
  async checkInstalled() {
    try {
      log.info('[Windows] Checking if PowerPoint is installed');
      const response = await bridge.sendCommand({ action: 'check' });
      
      if (response.status === 'success') {
        log.info('[Windows] PowerPoint is installed, version:', response.data);
        return true;
      }
      
      log.error('[Windows] PowerPoint check failed:', response.error);
      return false;
    } catch (error) {
      log.error('[Windows] PowerPoint check error:', error);
      return false;
    }
  },

  async openPresentation(filePath: string) {
    try {
      log.info('[Windows] Opening presentation:', filePath);
      currentPresentationPath = filePath;
      
      // Copy file to local temp directory to avoid permission issues with cloud storage
      const tempDir = join(tmpdir(), 'slidecue-presentations');
      await mkdir(tempDir, { recursive: true });
      localPresentationCopy = join(tempDir, basename(filePath));
      
      log.info('[Windows] Copying presentation to temp location:', localPresentationCopy);
      await copyFile(filePath, localPresentationCopy);
      
      // Open in PowerPoint via PowerShell bridge
      log.info('[Windows] Opening presentation in PowerPoint');
      const openResponse = await bridge.sendCommand({
        action: 'open',
        filePath: localPresentationCopy
      });
      
      if (openResponse.status !== 'success') {
        throw new Error(openResponse.error || 'Failed to open presentation');
      }
      
      // Get slide metadata from PowerPoint (including hidden slides and animations)
      log.info('[Windows] Getting slide metadata from PowerPoint');
      const metadataResponse = await bridge.sendCommand({ action: 'getMetadata' });
      
      if (metadataResponse.status === 'success' && metadataResponse.data) {
        const metadata = JSON.parse(metadataResponse.data);
        totalSlides = metadata.totalSlides;
        
        // Build presentationData from PowerPoint metadata
        const slides = metadata.slides.map((s: any) => ({
          slideNumber: s.slideNumber,
          name: `Slide ${s.slideNumber}`,
          hidden: s.hidden,
          animationClicks: s.animationClicks,
          notes: s.notes || ''
        }));
        
        const visibleSlides = slides.filter((s: any) => !s.hidden).map((s: any) => s.slideNumber);
        const hiddenSlides = slides.filter((s: any) => s.hidden).map((s: any) => s.slideNumber);
        
        presentationData = {
          slides,
          totalSlides,
          visibleSlides,
          hiddenSlides
        };
        
        log.info('[Windows] Parsed presentation metadata:', {
          totalSlides: presentationData.totalSlides,
          visibleSlides: presentationData.visibleSlides,
          hiddenSlides: presentationData.hiddenSlides
        });
      } else {
        // Fallback to basic data if metadata fails
        totalSlides = parseInt(openResponse.data || '0', 10);
        presentationData = {
          slides: [],
          totalSlides,
          visibleSlides: Array.from({ length: totalSlides }, (_, i) => i + 1),
          hiddenSlides: []
        };
        log.warn('[Windows] Failed to get metadata, using fallback');
      }
      
      // Get total slides from response
      totalSlides = parseInt(openResponse.data || '0', 10) || presentationData.totalSlides;
      currentSlide = 1;
      currentAnimationStep = 0;
      
      log.info('[Windows] Opened presentation with', totalSlides, 'slides');
      log.info('[Windows] Visible slides:', presentationData.visibleSlides.join(', '));
      log.info('[Windows] Hidden slides:', presentationData.hiddenSlides.join(', '));
    } catch (error) {
      log.error('[Windows] Failed to open presentation:', error);
      throw error;
    }
  },

  async exportThumbnails(outputDir: string, onProgress?: ProgressCallback): Promise<SlideMetadata> {
    try {
      log.info('[Windows] Exporting thumbnails to:', outputDir);
      await mkdir(outputDir, { recursive: true });
      
      onProgress?.(1, totalSlides);
      
      // Export thumbnails via PowerShell bridge
      const response = await bridge.sendCommand({
        action: 'export',
        outputDir
      }, 60000); // 60 second timeout for export
      
      if (response.status !== 'success') {
        throw new Error(response.error || 'Failed to export thumbnails');
      }
      
      const exportedCount = parseInt(response.data || '0', 10);
      log.info('[Windows] Exported', exportedCount, 'thumbnails');
      
      // Get paths to exported files
      const files = await readdir(outputDir);
      const paths = files
        .filter(f => f.endsWith('.png'))
        .sort()
        .map(f => join(outputDir, f));
      
      log.info('[Windows] Generated', paths.length, 'thumbnail files');
      
      return {
        thumbnails: paths,
        totalSlides,
        hiddenSlides: presentationData?.hiddenSlides || [],
        visibleSlides: presentationData?.visibleSlides || [],
      };
    } catch (error) {
      log.error('[Windows] Failed to export thumbnails:', error);
      throw error;
    }
  },

  async startSlideshow() {
    try {
      log.info('[Windows] Starting slideshow');
      currentSlide = 1;
      currentAnimationStep = 0;
      
      const response = await bridge.sendCommand({ action: 'start' });
      
      if (response.status !== 'success') {
        throw new Error(response.error || 'Failed to start slideshow');
      }
      
      // Get actual slide position from response
      currentSlide = parseInt(response.data || '1', 10);
      log.info('[Windows] Slideshow started at slide', currentSlide);
    } catch (error) {
      log.error('[Windows] Failed to start slideshow:', error);
      throw error;
    }
  },

  async nextSlide() {
    try {
      const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
      const animationsOnSlide = slideData?.animationClicks || 0;
      
      log.info('[Windows] Next slide (current:', currentSlide, 'animation:', currentAnimationStep, '/', animationsOnSlide, ')');
      
      const response = await bridge.sendCommand({ action: 'next' });
      
      if (response.status !== 'success') {
        log.error('[Windows] Failed to advance:', response.error);
        return;
      }
      
      const newSlide = parseInt(response.data || String(currentSlide), 10);
      
      // Check if we advanced an animation or moved to next slide
      if (newSlide !== currentSlide) {
        currentSlide = newSlide;
        currentAnimationStep = 0;
        log.info('[Windows] Moved to slide', currentSlide);
      } else if (currentAnimationStep < animationsOnSlide) {
        currentAnimationStep++;
        log.info('[Windows] Animation', currentAnimationStep, '/', animationsOnSlide, 'on slide', currentSlide);
      }
    } catch (error) {
      log.error('[Windows] Failed to advance slide:', error);
    }
  },

  async prevSlide() {
    try {
      log.info('[Windows] Previous slide (current:', currentSlide, ')');
      
      const response = await bridge.sendCommand({ action: 'previous' });
      
      if (response.status !== 'success') {
        log.error('[Windows] Failed to go back:', response.error);
        return;
      }
      
      currentSlide = parseInt(response.data || String(currentSlide), 10);
      currentAnimationStep = 0;
      log.info('[Windows] Moved to slide', currentSlide);
    } catch (error) {
      log.error('[Windows] Failed to go to previous slide:', error);
    }
  },

  async gotoSlide(slideNumber: number) {
    try {
      log.info('[Windows] Going to slide', slideNumber);
      
      const response = await bridge.sendCommand({
        action: 'goto',
        slideNumber
      });
      
      if (response.status !== 'success') {
        log.error('[Windows] Failed to go to slide:', response.error);
        return;
      }
      
      currentSlide = parseInt(response.data || String(slideNumber), 10);
      currentAnimationStep = 0;
      log.info('[Windows] Now on slide', currentSlide);
    } catch (error) {
      log.error('[Windows] Failed to go to slide:', error);
    }
  },

  async getSlideInfo(): Promise<SlideInfo> {
    try {
      // Query current slide position from PowerPoint
      const response = await bridge.sendCommand({ action: 'getCurrentSlide' });
      
      if (response.status === 'success') {
        const actualSlide = parseInt(response.data || String(currentSlide), 10);
        if (actualSlide !== currentSlide) {
          currentSlide = actualSlide;
          currentAnimationStep = 0;
        }
      }
    } catch (error) {
      log.error('[Windows] Failed to query current slide:', error);
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
      log.info('[Windows] Stopping slideshow');
      const response = await bridge.sendCommand({ action: 'stop' });
      
      if (response.status !== 'success') {
        log.error('[Windows] Failed to stop slideshow:', response.error);
      }
    } catch (error) {
      log.error('[Windows] Failed to stop slideshow:', error);
    }
  },

  async closePresentation() {
    try {
      log.info('[Windows] Closing presentation');
      
      // Close presentation via PowerShell bridge
      const response = await bridge.sendCommand({ action: 'close' });
      
      if (response.status !== 'success') {
        log.error('[Windows] Failed to close presentation:', response.error);
      }
      
      // Stop the bridge
      await bridge.stop();
      
      // Clean up temp copy
      if (localPresentationCopy) {
        try {
          await unlink(localPresentationCopy);
          log.info('[Windows] Deleted temp presentation copy');
        } catch (error) {
          log.error('[Windows] Failed to delete temp copy:', error);
        }
        localPresentationCopy = '';
      }
      
      // Reset state
      presentationData = null;
      currentSlide = 1;
      currentAnimationStep = 0;
      totalSlides = 1;
      currentPresentationPath = '';
    } catch (error) {
      log.error('[Windows] Failed to close presentation:', error);
    }
  },
};
