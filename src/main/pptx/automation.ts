import { PowerPointAutomation } from './types';
import { macOSAutomation } from './macos';
import { windowsAutomation } from './windows';

export function getAutomation(): PowerPointAutomation {
  if (process.platform === 'darwin') {
    return macOSAutomation;
  } else if (process.platform === 'win32') {
    return windowsAutomation;
  }
  throw new Error('Unsupported platform');
}

export * from './types';
