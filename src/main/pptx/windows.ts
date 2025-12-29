import log from 'electron-log';
import { PowerPointAutomation } from './types';
import { windowsWinaxAutomation } from './windows-winax';

// Always import PowerShell bridge as fallback (static import, not dynamic)
import { windowsAutomation as powershellAutomation } from './windows-powershell';

// Determine which automation to use
let automation: PowerPointAutomation;

// Try winax first
try {
  const winax = require('winax');
  automation = windowsWinaxAutomation;
  log.info('[Windows] Using winax for PowerPoint automation');
} catch (error) {
  log.warn('[Windows] winax not available, falling back to PowerShell bridge:', error);
  automation = powershellAutomation;
  log.info('[Windows] Using PowerShell bridge for PowerPoint automation');
}

export const windowsAutomation: PowerPointAutomation = automation;
