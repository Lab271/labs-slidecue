// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import log from 'electron-log';
import { windowsAutomation as powershellAutomation } from './windows-powershell';

log.info('[Windows] Using PowerShell bridge for PowerPoint automation');

export const windowsAutomation = powershellAutomation;
