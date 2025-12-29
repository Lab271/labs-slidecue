import log from 'electron-log';
import { windowsAutomation as powershellAutomation } from './windows-powershell';

log.info('[Windows] Using PowerShell bridge for PowerPoint automation');

export const windowsAutomation = powershellAutomation;
