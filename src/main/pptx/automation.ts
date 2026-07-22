// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
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
