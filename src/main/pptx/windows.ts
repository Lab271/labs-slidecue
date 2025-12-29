import log from 'electron-log';
import { PowerPointAutomation } from './types';
import { windowsWinaxAutomation } from './windows-winax';

// Try to load PowerShell bridge as fallback
let powershellAutomation: PowerPointAutomation | null = null;
try {
  const psModule = require('./windows-powershell');
  powershellAutomation = psModule.windowsAutomation;
} catch (error) {
  log.error('[Windows] Failed to load PowerShell fallback:', error);
}

// Determine which automation to use
let useWinax = false;
let automation: PowerPointAutomation;

// Try winax first
try {
  const winax = require('winax');
  useWinax = true;
  automation = windowsWinaxAutomation;
  log.info('[Windows] Using winax for PowerPoint automation');
} catch (error) {
  log.warn('[Windows] winax not available, falling back to PowerShell bridge:', error);
  if (powershellAutomation) {
    automation = powershellAutomation;
    log.info('[Windows] Using PowerShell bridge for PowerPoint automation');
  } else {
    throw new Error('Neither winax nor PowerShell bridge is available');
  }
}

export const windowsAutomation: PowerPointAutomation = automation;
